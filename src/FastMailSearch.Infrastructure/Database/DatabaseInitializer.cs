using System;
using Microsoft.Data.Sqlite;
using FastMailSearch.Infrastructure.Logging;

namespace FastMailSearch.Infrastructure.Database
{
    public class DatabaseInitializer
    {
        private readonly DatabaseConnectionFactory _connectionFactory;

        public DatabaseInitializer(DatabaseConnectionFactory connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public void Initialize()
        {
            try
            {
                using (var conn = _connectionFactory.CreateConnection())
                {
                    // Fast check: If core tables and unique index already exist, skip DDL and migrations completely (<1ms fast path)
                    using (var checkCmd = conn.CreateCommand())
                    {
                        checkCmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE (type='table' AND name IN ('Emails', 'IndexState', 'FailedIndexItems', 'EmailSearch', 'PendingChanges')) OR (type='index' AND name='UQ_FailedIndexItems_StoreId_EntryId');";
                        var countObj = checkCmd.ExecuteScalar();
                        int matchCount = countObj != null ? Convert.ToInt32(countObj) : 0;
                        if (matchCount >= 6)
                        {
                            Logger.Info("Database schema verified (tables and indexes already initialized).");
                            return;
                        }
                    }

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
                        CREATE TABLE IF NOT EXISTS Emails (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            StoreId TEXT NOT NULL,
                            EntryId TEXT NOT NULL,
                            ConversationId TEXT,
                            Subject TEXT,
                            SenderName TEXT,
                            SenderEmail TEXT,
                            Recipients TEXT,
                            Cc TEXT,
                            Bcc TEXT,
                            BodyText TEXT,
                            ReceivedDate TEXT,
                            SentDate TEXT,
                            FolderName TEXT,
                            FolderEntryId TEXT,
                            HasAttachment INTEGER NOT NULL DEFAULT 0,
                            IsRead INTEGER NOT NULL DEFAULT 0,
                            IsFlagged INTEGER NOT NULL DEFAULT 0,
                            Importance INTEGER NOT NULL DEFAULT 1,
                            Categories TEXT,
                            Size INTEGER NOT NULL DEFAULT 0,
                            LastModified TEXT,
                            IndexedDate TEXT NOT NULL,
                            UNIQUE(StoreId, EntryId)
                        );

                        CREATE INDEX IF NOT EXISTS IX_Emails_ReceivedDate ON Emails(ReceivedDate);
                        CREATE INDEX IF NOT EXISTS IX_Emails_FolderName ON Emails(FolderName);
                        CREATE INDEX IF NOT EXISTS IX_Emails_StoreId_FolderEntryId ON Emails(StoreId, FolderEntryId);
                        CREATE INDEX IF NOT EXISTS IX_Emails_SenderEmail ON Emails(SenderEmail);
                        CREATE INDEX IF NOT EXISTS IX_Emails_HasAttachment ON Emails(HasAttachment);
                        CREATE INDEX IF NOT EXISTS IX_Emails_Recipients ON Emails(Recipients);

                        CREATE TABLE IF NOT EXISTS Attachments (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            EmailId INTEGER NOT NULL,
                            Name TEXT NOT NULL,
                            Extension TEXT,
                            Size INTEGER NOT NULL DEFAULT 0,
                            FOREIGN KEY(EmailId) REFERENCES Emails(Id) ON DELETE CASCADE
                        );

                        CREATE INDEX IF NOT EXISTS IX_Attachments_EmailId ON Attachments(EmailId);

                        CREATE TABLE IF NOT EXISTS Folders (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            StoreId TEXT NOT NULL,
                            EntryId TEXT NOT NULL,
                            Name TEXT NOT NULL,
                            ParentId TEXT,
                            UNIQUE(StoreId, EntryId)
                        );

                        CREATE TABLE IF NOT EXISTS IndexState (
                            StoreId TEXT NOT NULL,
                            FolderId TEXT NOT NULL,
                            LastSync TEXT,
                            Status TEXT,
                            ItemCount INTEGER DEFAULT 0,
                            ProcessedOffset INTEGER DEFAULT 0,
                            PRIMARY KEY(StoreId, FolderId)
                        );

                        CREATE TABLE IF NOT EXISTS SearchHistory (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            QueryText TEXT NOT NULL,
                            ExecutedAt TEXT NOT NULL
                        );

                        CREATE TABLE IF NOT EXISTS GlobalState (
                            Key TEXT PRIMARY KEY,
                            Value TEXT NOT NULL,
                            UpdatedAt TEXT NOT NULL
                        );

                        CREATE TABLE IF NOT EXISTS PendingChanges (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            StoreId TEXT NOT NULL,
                            EntryId TEXT NOT NULL,
                            Operation TEXT NOT NULL,
                            FolderEntryId TEXT,
                            CreatedAt TEXT NOT NULL,
                            UpdatedAt TEXT NOT NULL,
                            AttemptCount INTEGER NOT NULL DEFAULT 0,
                            Status TEXT NOT NULL,
                            LastError TEXT
                        );

                        CREATE INDEX IF NOT EXISTS IX_PendingChanges_Status ON PendingChanges(Status);
                        CREATE INDEX IF NOT EXISTS IX_PendingChanges_StoreId_EntryId ON PendingChanges(StoreId, EntryId);

                        CREATE TABLE IF NOT EXISTS FailedIndexItems (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            StoreId TEXT NOT NULL,
                            EntryId TEXT NOT NULL,
                            FolderName TEXT,
                            OperationType TEXT NOT NULL,
                            FailureReason TEXT,
                            ExceptionType TEXT,
                            AttemptCount INTEGER NOT NULL DEFAULT 0,
                            FirstFailedAt TEXT NOT NULL,
                            LastFailedAt TEXT NOT NULL,
                            LastRetryAt TEXT,
                            Status TEXT NOT NULL DEFAULT 'Failed',
                            ErrorMessage TEXT,
                            Resolved INTEGER NOT NULL DEFAULT 0,
                            UNIQUE(StoreId, EntryId)
                        );

                        CREATE INDEX IF NOT EXISTS IX_FailedIndexItems_StoreId_EntryId ON FailedIndexItems(StoreId, EntryId);
                        CREATE INDEX IF NOT EXISTS IX_FailedIndexItems_Status ON FailedIndexItems(Status);
                        CREATE INDEX IF NOT EXISTS IX_FailedIndexItems_LastFailedAt ON FailedIndexItems(LastFailedAt);
                        CREATE INDEX IF NOT EXISTS IX_FailedIndexItems_OperationType ON FailedIndexItems(OperationType);

                        CREATE VIRTUAL TABLE IF NOT EXISTS EmailSearch USING fts5(
                            EmailId UNINDEXED,
                            Subject,
                            SenderName,
                            SenderEmail,
                            Recipients,
                            BodyText,
                            AttachmentNames,
                            tokenize = 'unicode61'
                        );
                    ";

                        cmd.ExecuteNonQuery();

                        // Migration: Ensure FailedIndexItems has no duplicates and has a unique index
                        MigrateFailedIndexItemsUniqueConstraint(conn);

                        // Migration: Ensure IndexState has ProcessedOffset column for resumable backfill
                        MigrateIndexStateProcessedOffset(conn);

                        Logger.Info("Database tables and FTS5 index initialized successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to initialize database schema", ex);
                throw;
            }
        }

        private void MigrateIndexStateProcessedOffset(SqliteConnection conn)
        {
            try
            {
                using (var checkCmd = conn.CreateCommand())
                {
                    checkCmd.CommandText = "PRAGMA table_info(IndexState);";
                    using (var reader = checkCmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (string.Equals(reader.GetString(1), "ProcessedOffset", StringComparison.OrdinalIgnoreCase))
                            {
                                return; // Column already exists, safe to skip
                            }
                        }
                    }
                }

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "ALTER TABLE IndexState ADD COLUMN ProcessedOffset INTEGER DEFAULT 0;";
                    cmd.ExecuteNonQuery();
                }
            }
            catch (SqliteException)
            {
                // Column already exists, safe to ignore
            }
            catch (Exception ex)
            {
                Logger.Warn($"Notice during IndexState migration: {ex.Message}");
            }
        }

        private void MigrateFailedIndexItemsUniqueConstraint(SqliteConnection conn)
        {
            try
            {
                using (var checkCmd = conn.CreateCommand())
                {
                    checkCmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name='UQ_FailedIndexItems_StoreId_EntryId';";
                    var res = checkCmd.ExecuteScalar();
                    if (res != null && Convert.ToInt32(res) > 0)
                    {
                        return; // Unique index already exists, skip expensive migration transaction
                    }
                }

                using (var tx = conn.BeginTransaction())
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
                            -- Create a temporary table with the consolidated data
                            CREATE TEMP TABLE IF NOT EXISTS ConsolidatedFailedItems AS
                            SELECT 
                                MAX(Id) as KeepId,
                                StoreId, 
                                EntryId, 
                                MIN(FirstFailedAt) as MinFirstFailedAt,
                                SUM(AttemptCount) as TotalAttemptCount
                            FROM FailedIndexItems
                            GROUP BY StoreId, EntryId
                            HAVING COUNT(*) > 1;

                            -- Update the kept records with consolidated data
                            UPDATE FailedIndexItems
                            SET 
                                FirstFailedAt = (SELECT MinFirstFailedAt FROM ConsolidatedFailedItems WHERE KeepId = FailedIndexItems.Id),
                                AttemptCount = (SELECT TotalAttemptCount FROM ConsolidatedFailedItems WHERE KeepId = FailedIndexItems.Id)
                            WHERE Id IN (SELECT KeepId FROM ConsolidatedFailedItems);

                            -- Delete the duplicates
                            DELETE FROM FailedIndexItems
                            WHERE Id NOT IN (
                                SELECT MAX(Id) 
                                FROM FailedIndexItems 
                                GROUP BY StoreId, EntryId
                            );

                            DROP TABLE IF EXISTS ConsolidatedFailedItems;
                        ";
                        cmd.ExecuteNonQuery();
                    }

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = "CREATE UNIQUE INDEX IF NOT EXISTS UQ_FailedIndexItems_StoreId_EntryId ON FailedIndexItems(StoreId, EntryId);";
                        cmd.ExecuteNonQuery();
                    }

                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to migrate FailedIndexItems to enforce unique constraint", ex);
                throw;
            }
        }
    }
}
