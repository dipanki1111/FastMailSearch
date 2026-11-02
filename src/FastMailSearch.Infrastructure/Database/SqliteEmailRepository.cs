using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Search;
using FastMailSearch.Infrastructure.Logging;
using Microsoft.Data.Sqlite;

namespace FastMailSearch.Infrastructure.Database
{
    public class SqliteEmailRepository : IEmailRepository, ISearchEngine
    {
        private readonly DatabaseConnectionFactory _factory;
        private readonly DatabaseInitializer _initializer;
        private readonly object _folderLock = new object();
        private List<string>? _cachedFolderNames;

        public SqliteEmailRepository(DatabaseConnectionFactory factory)
        {
            _factory = factory ?? throw new ArgumentNullException(nameof(factory));
            _initializer = new DatabaseInitializer(_factory);
        }

        public string DatabasePath => _factory.DbPath;
        public DatabaseConnectionFactory ConnectionFactory => _factory;

        public void InitializeDatabase()
        {
            _initializer.Initialize();
            lock (_folderLock) { _cachedFolderNames = null; }
        }

        public Task SaveEmailBatchAsync(IEnumerable<EmailDocument> emails, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                if (emails == null) return;
                var list = emails.ToList();
                if (list.Count == 0) return;

                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (var email in list)
                        {
                            if (cancellationToken.IsCancellationRequested)
                                break;

                            long emailId = UpsertEmail(conn, email, tx);
                            email.Id = emailId;

                            // Update attachments
                            using (var delAtt = conn.CreateCommand())
                            {
                                delAtt.Transaction = tx;
                                delAtt.CommandText = "DELETE FROM Attachments WHERE EmailId = @EmailId;";
                                delAtt.Parameters.AddWithValue("@EmailId", emailId);
                                delAtt.ExecuteNonQuery();
                            }

                            if (email.Attachments != null && email.Attachments.Count > 0)
                            {
                                foreach (var att in email.Attachments)
                                {
                                    using (var insAtt = conn.CreateCommand())
                                    {
                                        insAtt.Transaction = tx;
                                        insAtt.CommandText = @"
                                            INSERT INTO Attachments (EmailId, Name, Extension, Size)
                                            VALUES (@EmailId, @Name, @Extension, @Size);";
                                        insAtt.Parameters.AddWithValue("@EmailId", emailId);
                                        insAtt.Parameters.AddWithValue("@Name", att.Name ?? string.Empty);
                                        insAtt.Parameters.AddWithValue("@Extension", att.Extension ?? string.Empty);
                                        insAtt.Parameters.AddWithValue("@Size", att.Size);
                                        insAtt.ExecuteNonQuery();
                                    }
                                }
                            }

                            // Upsert FTS table
                            using (var delFts = conn.CreateCommand())
                            {
                                delFts.Transaction = tx;
                                delFts.CommandText = "DELETE FROM EmailSearch WHERE EmailId = @EmailId;";
                                delFts.Parameters.AddWithValue("@EmailId", emailId.ToString(CultureInfo.InvariantCulture));
                                delFts.ExecuteNonQuery();
                            }

                            using (var insFts = conn.CreateCommand())
                            {
                                insFts.Transaction = tx;
                                insFts.CommandText = @"
                                    INSERT INTO EmailSearch (EmailId, Subject, SenderName, SenderEmail, Recipients, BodyText, AttachmentNames)
                                    VALUES (@EmailId, @Subject, @SenderName, @SenderEmail, @Recipients, @BodyText, @AttachmentNames);";
                                insFts.Parameters.AddWithValue("@EmailId", emailId.ToString(CultureInfo.InvariantCulture));
                                insFts.Parameters.AddWithValue("@Subject", email.Subject ?? string.Empty);
                                insFts.Parameters.AddWithValue("@SenderName", email.SenderName ?? string.Empty);
                                insFts.Parameters.AddWithValue("@SenderEmail", email.SenderEmail ?? string.Empty);
                                insFts.Parameters.AddWithValue("@Recipients", email.Recipients ?? string.Empty);
                                insFts.Parameters.AddWithValue("@BodyText", email.BodyText ?? string.Empty);
                                insFts.Parameters.AddWithValue("@AttachmentNames", email.GetAttachmentNamesJoined());
                                insFts.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error("Failed saving email batch transaction", ex);
                        throw;
                    }
                }
            }, cancellationToken);
        }

        private long UpsertEmail(SqliteConnection conn, EmailDocument email, SqliteTransaction tx)
        {
            long existingId = 0;
            using (var check = conn.CreateCommand())
            {
                check.Transaction = tx;
                check.CommandText = "SELECT Id FROM Emails WHERE StoreId = @StoreId AND EntryId = @EntryId;";
                check.Parameters.AddWithValue("@StoreId", email.StoreId);
                check.Parameters.AddWithValue("@EntryId", email.EntryId);
                object? result = check.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    existingId = Convert.ToInt64(result);
                }
            }

            string nowStr = DateTime.Now.ToString("o");

            if (existingId > 0)
            {
                using (var update = conn.CreateCommand())
                {
                    update.Transaction = tx;
                    update.CommandText = @"
                        UPDATE Emails SET
                            ConversationId = @ConversationId,
                            Subject = @Subject,
                            SenderName = @SenderName,
                            SenderEmail = @SenderEmail,
                            Recipients = @Recipients,
                            Cc = @Cc,
                            Bcc = @Bcc,
                            BodyText = @BodyText,
                            ReceivedDate = @ReceivedDate,
                            SentDate = @SentDate,
                            FolderName = @FolderName,
                            FolderEntryId = @FolderEntryId,
                            HasAttachment = @HasAttachment,
                            IsRead = @IsRead,
                            IsFlagged = @IsFlagged,
                            Importance = @Importance,
                            Categories = @Categories,
                            Size = @Size,
                            LastModified = @LastModified,
                            IndexedDate = @IndexedDate
                        WHERE Id = @Id;";

                    AddEmailParameters(update, email, nowStr);
                    update.Parameters.AddWithValue("@Id", existingId);
                    update.ExecuteNonQuery();
                }
                return existingId;
            }
            else
            {
                using (var insert = conn.CreateCommand())
                {
                    insert.Transaction = tx;
                    insert.CommandText = @"
                        INSERT INTO Emails (
                            StoreId, EntryId, ConversationId, Subject, SenderName, SenderEmail,
                            Recipients, Cc, Bcc, BodyText, ReceivedDate, SentDate, FolderName,
                            FolderEntryId, HasAttachment, IsRead, IsFlagged, Importance,
                            Categories, Size, LastModified, IndexedDate
                        ) VALUES (
                            @StoreId, @EntryId, @ConversationId, @Subject, @SenderName, @SenderEmail,
                            @Recipients, @Cc, @Bcc, @BodyText, @ReceivedDate, @SentDate, @FolderName,
                            @FolderEntryId, @HasAttachment, @IsRead, @IsFlagged, @Importance,
                            @Categories, @Size, @LastModified, @IndexedDate
                        );
                        SELECT last_insert_rowid();";

                    AddEmailParameters(insert, email, nowStr);
                    long newId = Convert.ToInt64(insert.ExecuteScalar());
                    return newId;
                }
            }
        }

        private void AddEmailParameters(SqliteCommand cmd, EmailDocument email, string nowStr)
        {
            cmd.Parameters.AddWithValue("@StoreId", email.StoreId ?? string.Empty);
            cmd.Parameters.AddWithValue("@EntryId", email.EntryId ?? string.Empty);
            cmd.Parameters.AddWithValue("@ConversationId", email.ConversationId ?? string.Empty);
            cmd.Parameters.AddWithValue("@Subject", email.Subject ?? string.Empty);
            cmd.Parameters.AddWithValue("@SenderName", email.SenderName ?? string.Empty);
            cmd.Parameters.AddWithValue("@SenderEmail", email.SenderEmail ?? string.Empty);
            cmd.Parameters.AddWithValue("@Recipients", email.Recipients ?? string.Empty);
            cmd.Parameters.AddWithValue("@Cc", email.Cc ?? string.Empty);
            cmd.Parameters.AddWithValue("@Bcc", email.Bcc ?? string.Empty);
            cmd.Parameters.AddWithValue("@BodyText", email.BodyText ?? string.Empty);
            cmd.Parameters.AddWithValue("@ReceivedDate", email.ReceivedDate.ToString("o"));
            cmd.Parameters.AddWithValue("@SentDate", email.SentDate.ToString("o"));
            cmd.Parameters.AddWithValue("@FolderName", email.FolderName ?? string.Empty);
            cmd.Parameters.AddWithValue("@FolderEntryId", email.FolderEntryId ?? string.Empty);
            cmd.Parameters.AddWithValue("@HasAttachment", email.HasAttachment ? 1 : 0);
            cmd.Parameters.AddWithValue("@IsRead", email.IsRead ? 1 : 0);
            cmd.Parameters.AddWithValue("@IsFlagged", email.IsFlagged ? 1 : 0);
            cmd.Parameters.AddWithValue("@Importance", email.Importance);
            cmd.Parameters.AddWithValue("@Categories", email.Categories ?? string.Empty);
            cmd.Parameters.AddWithValue("@Size", email.Size);
            cmd.Parameters.AddWithValue("@LastModified", email.LastModified.ToString("o"));
            cmd.Parameters.AddWithValue("@IndexedDate", nowStr);
        }

        public Task<EmailDocument?> GetEmailByStoreAndEntryIdAsync(string storeId, string entryId)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM Emails WHERE StoreId = @StoreId AND EntryId = @EntryId;";
                    cmd.Parameters.AddWithValue("@StoreId", storeId);
                    cmd.Parameters.AddWithValue("@EntryId", entryId);

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            var email = ReadEmailDocument(reader);
                            email.Attachments = GetEmailAttachments(conn, email.Id);
                            return (EmailDocument?)email;
                        }
                    }
                }
                return null;
            });
        }

        public Task DeleteEmailAsync(string storeId, string entryId)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        long emailId = 0;
                        using (var getId = conn.CreateCommand())
                        {
                            getId.Transaction = tx;
                            getId.CommandText = "SELECT Id FROM Emails WHERE StoreId = @StoreId AND EntryId = @EntryId;";
                            getId.Parameters.AddWithValue("@StoreId", storeId);
                            getId.Parameters.AddWithValue("@EntryId", entryId);
                            object? res = getId.ExecuteScalar();
                            if (res != null && res != DBNull.Value) emailId = Convert.ToInt64(res);
                        }

                        if (emailId > 0)
                        {
                            using (var delE = conn.CreateCommand())
                            {
                                delE.Transaction = tx;
                                delE.CommandText = "DELETE FROM Emails WHERE Id = @Id;";
                                delE.Parameters.AddWithValue("@Id", emailId);
                                delE.ExecuteNonQuery();
                            }

                            using (var delF = conn.CreateCommand())
                            {
                                delF.Transaction = tx;
                                delF.CommandText = "DELETE FROM EmailSearch WHERE EmailId = @EmailId;";
                                delF.Parameters.AddWithValue("@EmailId", emailId.ToString(CultureInfo.InvariantCulture));
                                delF.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error($"Error deleting email ({storeId}, {entryId})", ex);
                    }
                }
            });
        }

        public Task UpdateEmailFolderAsync(string storeId, string entryId, string newFolderName, string newFolderEntryId)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE Emails SET FolderName = @FolderName, FolderEntryId = @FolderEntryId WHERE StoreId = @StoreId AND EntryId = @EntryId;";
                    cmd.Parameters.AddWithValue("@FolderName", newFolderName);
                    cmd.Parameters.AddWithValue("@FolderEntryId", newFolderEntryId);
                    cmd.Parameters.AddWithValue("@StoreId", storeId);
                    cmd.Parameters.AddWithValue("@EntryId", entryId);
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task<HashSet<string>> GetIndexedEntryIdsForFolderAsync(string storeId, string folderEntryId)
        {
            return Task.Run(() =>
            {
                var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT EntryId FROM Emails WHERE StoreId = @StoreId AND FolderEntryId = @FolderEntryId;";
                    cmd.Parameters.AddWithValue("@StoreId", storeId);
                    cmd.Parameters.AddWithValue("@FolderEntryId", folderEntryId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            set.Add(reader.GetString(0));
                        }
                    }
                }
                return set;
            });
        }

        public Task SaveFoldersAsync(IEnumerable<MailboxFolder> folders)
        {
            return Task.Run(() =>
            {
                if (folders == null) return;
                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    foreach (var f in folders)
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            cmd.CommandText = @"
                                INSERT INTO Folders (StoreId, EntryId, Name, ParentId)
                                VALUES (@StoreId, @EntryId, @Name, @ParentId)
                                ON CONFLICT(StoreId, EntryId) DO UPDATE SET
                                Name = excluded.Name, ParentId = excluded.ParentId;";
                            cmd.Parameters.AddWithValue("@StoreId", f.StoreId);
                            cmd.Parameters.AddWithValue("@EntryId", f.EntryId);
                            cmd.Parameters.AddWithValue("@Name", f.Name);
                            cmd.Parameters.AddWithValue("@ParentId", f.ParentId ?? string.Empty);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    tx.Commit();
                }
            });
        }

        public Task<List<MailboxFolder>> GetFoldersAsync(string? storeId = null)
        {
            return Task.Run(() =>
            {
                var list = new List<MailboxFolder>();
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    if (string.IsNullOrEmpty(storeId))
                    {
                        cmd.CommandText = "SELECT Id, StoreId, EntryId, Name, ParentId FROM Folders ORDER BY Name;";
                    }
                    else
                    {
                        cmd.CommandText = "SELECT Id, StoreId, EntryId, Name, ParentId FROM Folders WHERE StoreId = @StoreId ORDER BY Name;";
                        cmd.Parameters.AddWithValue("@StoreId", storeId);
                    }

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(new MailboxFolder
                            {
                                Id = reader.GetInt64(0),
                                StoreId = reader.GetString(1),
                                EntryId = reader.GetString(2),
                                Name = reader.GetString(3),
                                ParentId = reader.IsDBNull(4) ? "" : reader.GetString(4)
                            });
                        }
                    }
                }
                return list;
            });
        }

        public Task<IndexMetadata> GetIndexMetadataAsync()
        {
            return Task.Run(() =>
            {
                var meta = new IndexMetadata();
                using (var conn = _factory.CreateConnection())
                {
                    using (var cmdCount = conn.CreateCommand())
                    {
                        cmdCount.CommandText = "SELECT COUNT(*) FROM Emails;";
                        object? res = cmdCount.ExecuteScalar();
                        meta.TotalEmailsIndexed = res != null ? Convert.ToInt32(res) : 0;
                    }

                    using (var cmdLast = conn.CreateCommand())
                    {
                        cmdLast.CommandText = "SELECT MAX(IndexedDate) FROM Emails;";
                        object? res = cmdLast.ExecuteScalar();
                        if (res != null && res != DBNull.Value && DateTime.TryParse(res.ToString(), out DateTime lastDt))
                        {
                            meta.LastIndexedTime = lastDt;
                        }
                    }
                }
                return meta;
            });
        }

        public Task ClearAllDataAsync()
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
                            DELETE FROM Attachments;
                            DELETE FROM Emails;
                            DELETE FROM Folders;
                            DELETE FROM IndexState;
                            DELETE FROM EmailSearch;";
                        cmd.ExecuteNonQuery();
                    }
                    tx.Commit();
                }
                Logger.Info("Database index data cleared.");
            });
        }

        public Task SaveSearchQueryAsync(string rawQuery)
        {
            return Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(rawQuery)) return;
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO SearchHistory (QueryText, ExecutedAt) VALUES (@QueryText, @ExecutedAt);";
                    cmd.Parameters.AddWithValue("@QueryText", rawQuery.Trim());
                    cmd.Parameters.AddWithValue("@ExecutedAt", DateTime.Now.ToString("o"));
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task<List<string>> GetRecentSearchQueriesAsync(int count = 20)
        {
            return Task.Run(() =>
            {
                var queries = new List<string>();
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT DISTINCT QueryText FROM SearchHistory ORDER BY Id DESC LIMIT @Count;";
                    cmd.Parameters.AddWithValue("@Count", count);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            queries.Add(reader.GetString(0));
                        }
                    }
                }
                return queries;
            });
        }

        public Task ClearSearchHistoryAsync()
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "DELETE FROM SearchHistory;";
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task VacuumDatabaseAsync()
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "VACUUM;";
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task SaveIndexStateAsync(string state, int processed, int total, string currentFolder, string? error)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        string nowStr = DateTime.Now.ToString("o");
                        UpsertGlobalStateKey(conn, tx, "IndexState", state, nowStr);
                        UpsertGlobalStateKey(conn, tx, "LastProcessedCount", processed.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "LastTotalCount", total.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "LastIndexedFolder", currentFolder ?? "", nowStr);
                        UpsertGlobalStateKey(conn, tx, "LastError", error ?? "", nowStr);
                        if (state == "Indexing") UpsertGlobalStateKey(conn, tx, "IndexStartedAt", nowStr, nowStr);
                        else if (state == "Completed") UpsertGlobalStateKey(conn, tx, "IndexCompletedAt", nowStr, nowStr);

                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error("Error saving index state to database", ex);
                    }
                }
            });
        }

        private void UpsertGlobalStateKey(SqliteConnection conn, SqliteTransaction tx, string key, string val, string nowStr)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = @"
                    INSERT INTO GlobalState (Key, Value, UpdatedAt)
                    VALUES (@Key, @Value, @UpdatedAt)
                    ON CONFLICT(Key) DO UPDATE SET Value = excluded.Value, UpdatedAt = excluded.UpdatedAt;";
                cmd.Parameters.AddWithValue("@Key", key);
                cmd.Parameters.AddWithValue("@Value", val ?? "");
                cmd.Parameters.AddWithValue("@UpdatedAt", nowStr);
                cmd.ExecuteNonQuery();
            }
        }

        public Task<Dictionary<string, string>> GetPersistedIndexStateAsync()
        {
            return Task.Run(() =>
            {
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT Key, Value FROM GlobalState;";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            dict[reader.GetString(0)] = reader.GetString(1);
                        }
                    }
                }
                return dict;
            });
        }

        public Task<int> GetIndexedEmailCountAsync()
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM Emails;";
                    object? res = cmd.ExecuteScalar();
                    return res != null ? Convert.ToInt32(res) : 0;
                }
            });
        }

        public Task<Dictionary<string, int>> GetStoreEmailCountsAsync()
        {
            return Task.Run(() =>
            {
                var dict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT StoreId, COUNT(*) FROM Emails GROUP BY StoreId;";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string storeId = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
                            int count = reader.GetInt32(1);
                            dict[storeId] = count;
                        }
                    }
                }
                return dict;
            });
        }

        public Task<int> GetFolderEmailCountAsync(string storeId, string folderEntryId)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM Emails WHERE StoreId = @StoreId AND FolderEntryId = @FolderEntryId;";
                    cmd.Parameters.AddWithValue("@StoreId", storeId ?? string.Empty);
                    cmd.Parameters.AddWithValue("@FolderEntryId", folderEntryId ?? string.Empty);
                    object? res = cmd.ExecuteScalar();
                    return res != null ? Convert.ToInt32(res) : 0;
                }
            });
        }

        public Task SaveFolderIndexStateAsync(string storeId, string folderEntryId, string status, int processedOffset, int totalItems)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        string nowStr = DateTime.Now.ToString("o");
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            cmd.CommandText = @"
                                INSERT INTO IndexState (StoreId, FolderId, LastSync, Status, ItemCount, ProcessedOffset)
                                VALUES (@StoreId, @FolderId, @LastSync, @Status, @ItemCount, @ProcessedOffset)
                                ON CONFLICT(StoreId, FolderId) DO UPDATE SET
                                    LastSync = excluded.LastSync,
                                    Status = excluded.Status,
                                    ItemCount = excluded.ItemCount,
                                    ProcessedOffset = excluded.ProcessedOffset;";
                            cmd.Parameters.AddWithValue("@StoreId", storeId ?? string.Empty);
                            cmd.Parameters.AddWithValue("@FolderId", folderEntryId ?? string.Empty);
                            cmd.Parameters.AddWithValue("@LastSync", nowStr);
                            cmd.Parameters.AddWithValue("@Status", status ?? "Indexing");
                            cmd.Parameters.AddWithValue("@ItemCount", totalItems);
                            cmd.Parameters.AddWithValue("@ProcessedOffset", processedOffset);
                            cmd.ExecuteNonQuery();
                        }
                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error($"Error saving folder index state for {folderEntryId}", ex);
                    }
                }
            });
        }

        public Task<(string Status, int ProcessedOffset, int TotalItems)?> GetFolderIndexStateAsync(string storeId, string folderEntryId)
        {
            return Task.Run<(string Status, int ProcessedOffset, int TotalItems)?>(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT Status, ProcessedOffset, ItemCount FROM IndexState WHERE StoreId = @StoreId AND FolderId = @FolderId;";
                    cmd.Parameters.AddWithValue("@StoreId", storeId ?? string.Empty);
                    cmd.Parameters.AddWithValue("@FolderId", folderEntryId ?? string.Empty);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string status = reader.IsDBNull(0) ? "Pending" : reader.GetString(0);
                            int offset = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
                            int total = reader.IsDBNull(2) ? 0 : reader.GetInt32(2);
                            return (status, offset, total);
                        }
                    }
                }
                return null;
            });
        }

        public Task<List<FolderIndexCheckpoint>> GetAllFolderIndexStatesAsync()
        {
            return Task.Run(() =>
            {
                var list = new List<FolderIndexCheckpoint>();
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT StoreId, FolderId, Status, ProcessedOffset, ItemCount, LastSync FROM IndexState;";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(new FolderIndexCheckpoint
                            {
                                StoreId = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                                FolderId = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                                Status = reader.IsDBNull(2) ? "Pending" : reader.GetString(2),
                                ProcessedOffset = reader.IsDBNull(3) ? 0 : reader.GetInt32(3),
                                TotalItems = reader.IsDBNull(4) ? 0 : reader.GetInt32(4),
                                LastSync = reader.IsDBNull(5) ? string.Empty : reader.GetString(5)
                            });
                        }
                    }
                }
                return list;
            });
        }

        public Task SaveLiveProgressAsync(
            string state,
            string currentOperation,
            int processed,
            int total,
            string currentFolder,
            int folderOffset,
            int folderTotal,
            int newlyIndexed,
            int skipped,
            int failed,
            int pendingQueue)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        string nowStr = DateTime.Now.ToString("o");
                        UpsertGlobalStateKey(conn, tx, "IndexState", state, nowStr);
                        UpsertGlobalStateKey(conn, tx, "CurrentOperation", currentOperation ?? "", nowStr);
                        UpsertGlobalStateKey(conn, tx, "LastProcessedCount", processed.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "LastTotalCount", total.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "LastIndexedFolder", currentFolder ?? "", nowStr);
                        UpsertGlobalStateKey(conn, tx, "FolderProcessedOffset", folderOffset.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "FolderTotalItems", folderTotal.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "NewlyIndexedCount", newlyIndexed.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "SkippedCount", skipped.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "FailedCount", failed.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "PendingQueueCount", pendingQueue.ToString(), nowStr);
                        UpsertGlobalStateKey(conn, tx, "LastSync", nowStr, nowStr);

                        if (state == "Indexing") UpsertGlobalStateKey(conn, tx, "IndexStartedAt", nowStr, nowStr);
                        else if (state == "Completed") UpsertGlobalStateKey(conn, tx, "IndexCompletedAt", nowStr, nowStr);

                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error("Error saving live progress to database", ex);
                    }
                }
            });
        }

        // ISearchEngine Implementation
        public Task<SearchResultContainer> SearchAsync(SearchQuery query, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                var container = new SearchResultContainer();
                var sw = Stopwatch.StartNew();

                if (query.IsInvalidSyntax)
                {
                    Logger.Warn($"[SEARCH] Aborting search due to invalid syntax: '{query.RawQuery}' - {query.SyntaxErrorMessage}");
                    container.TotalCount = 0;
                    container.SearchTimeMs = sw.ElapsedMilliseconds;
                    return container;
                }

                using (var conn = _factory.CreateConnection())
                {
                    var whereClauses = new List<string>();
                    var sqlParams = new List<SqliteParameter>();

                    // Handle FTS Full-Text Search
                    string? ftsCondition = BuildFtsCondition(query, sqlParams);
                    bool hasFts = !string.IsNullOrEmpty(ftsCondition);

                    string fromClause;
                    if (hasFts)
                    {
                        fromClause = "FROM EmailSearch fts INNER JOIN Emails e ON CAST(fts.EmailId AS INTEGER) = e.Id";
                        whereClauses.Add(ftsCondition!);
                    }
                    else
                    {
                        fromClause = "FROM Emails e";
                    }

                    // Phase 7X: Safe emoji fallback for standalone or mixed emoji queries
                    var allEmojis = new List<string>();
                    allEmojis.AddRange(ExtractEmojis(query.FreeText));
                    foreach (var ep in query.ExactPhrases)
                    {
                        allEmojis.AddRange(ExtractEmojis(ep));
                    }
                    if (!string.IsNullOrWhiteSpace(query.SubjectFilter))
                    {
                        allEmojis.AddRange(ExtractEmojis(query.SubjectFilter));
                    }
                    var distinctEmojis = allEmojis.Distinct().ToList();
                    for (int ei = 0; ei < distinctEmojis.Count; ei++)
                    {
                        string pName = $"@EmojiFilter{ei}";
                        whereClauses.Add($"(e.Subject LIKE {pName} OR e.BodyText LIKE {pName})");
                        sqlParams.Add(new SqliteParameter(pName, $"%{distinctEmojis[ei]}%"));
                    }

                    // Field specific filters
                    if (!string.IsNullOrWhiteSpace(query.SenderFilter))
                    {
                        whereClauses.Add("(e.SenderEmail = @SenderExact OR e.SenderName LIKE @SenderFilter OR e.SenderEmail LIKE @SenderFilter)");
                        sqlParams.Add(new SqliteParameter("@SenderExact", query.SenderFilter));
                        sqlParams.Add(new SqliteParameter("@SenderFilter", $"%{query.SenderFilter}%"));
                    }

                    if (!string.IsNullOrWhiteSpace(query.RecipientFilter))
                    {
                        whereClauses.Add("(e.Recipients LIKE @RecipientFilter OR e.Cc LIKE @RecipientFilter)");
                        sqlParams.Add(new SqliteParameter("@RecipientFilter", $"%{query.RecipientFilter}%"));
                    }

                    if (!string.IsNullOrWhiteSpace(query.CcFilter))
                    {
                        whereClauses.Add("e.Cc LIKE @CcFilter");
                        sqlParams.Add(new SqliteParameter("@CcFilter", $"%{query.CcFilter}%"));
                    }

                    if (!string.IsNullOrWhiteSpace(query.SubjectFilter))
                    {
                        whereClauses.Add("e.Subject LIKE @SubjectFilter");
                        sqlParams.Add(new SqliteParameter("@SubjectFilter", $"%{query.SubjectFilter}%"));
                    }

                    if (!string.IsNullOrWhiteSpace(query.FolderFilter))
                    {
                        string rawFilter = query.FolderFilter.Trim();
                        if (rawFilter.Contains("%") || rawFilter.Contains("*"))
                        {
                            whereClauses.Add("e.FolderName LIKE @FolderFilter");
                            sqlParams.Add(new SqliteParameter("@FolderFilter", rawFilter.Replace("*", "%")));
                        }
                        else
                        {
                            var known = GetKnownFolderNames(conn);
                            var exactMatches = known.Where(k => string.Equals(k, rawFilter, StringComparison.OrdinalIgnoreCase)).ToList();
                            if (exactMatches.Count == 0)
                            {
                                exactMatches = known.Where(k => k.IndexOf(rawFilter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                            }

                            if (exactMatches.Count == 1)
                            {
                                whereClauses.Add("e.FolderName = @FolderFilter");
                                sqlParams.Add(new SqliteParameter("@FolderFilter", exactMatches[0]));
                            }
                            else if (exactMatches.Count > 1)
                            {
                                var folderParams = new List<string>();
                                for (int fi = 0; fi < exactMatches.Count; fi++)
                                {
                                    string pName = $"@FolderFilter{fi}";
                                    folderParams.Add(pName);
                                    sqlParams.Add(new SqliteParameter(pName, exactMatches[fi]));
                                }
                                whereClauses.Add($"e.FolderName IN ({string.Join(", ", folderParams)})");
                            }
                            else
                            {
                                whereClauses.Add("e.FolderName LIKE @FolderFilter");
                                sqlParams.Add(new SqliteParameter("@FolderFilter", $"%{rawFilter}%"));
                            }
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(query.StoreFilter))
                    {
                        whereClauses.Add("e.StoreId = @StoreFilter");
                        sqlParams.Add(new SqliteParameter("@StoreFilter", query.StoreFilter));
                    }

                    if (query.HasAttachmentFilter.HasValue)
                    {
                        whereClauses.Add("e.HasAttachment = @HasAttachmentFilter");
                        sqlParams.Add(new SqliteParameter("@HasAttachmentFilter", query.HasAttachmentFilter.Value ? 1 : 0));
                    }

                    if (!string.IsNullOrWhiteSpace(query.FilenameFilter))
                    {
                        whereClauses.Add("e.Id IN (SELECT EmailId FROM Attachments WHERE Name LIKE @FilenameFilter)");
                        sqlParams.Add(new SqliteParameter("@FilenameFilter", $"%{query.FilenameFilter}%"));
                    }

                    if (query.IsReadFilter.HasValue)
                    {
                        whereClauses.Add("e.IsRead = @IsReadFilter");
                        sqlParams.Add(new SqliteParameter("@IsReadFilter", query.IsReadFilter.Value ? 1 : 0));
                    }

                    if (query.IsFlaggedFilter.HasValue)
                    {
                        whereClauses.Add("e.IsFlagged = @IsFlaggedFilter");
                        sqlParams.Add(new SqliteParameter("@IsFlaggedFilter", query.IsFlaggedFilter.Value ? 1 : 0));
                    }

                    if (query.DateFrom.HasValue)
                    {
                        whereClauses.Add("e.ReceivedDate >= @DateFrom");
                        var dtFrom = query.DateFrom.Value;
                        string fromStr = dtFrom.ToString("yyyy-MM-ddTHH:mm:ss.fffffff");
                        sqlParams.Add(new SqliteParameter("@DateFrom", fromStr));
                    }

                    if (query.DateTo.HasValue)
                    {
                        whereClauses.Add("e.ReceivedDate <= @DateTo");
                        var dtTo = query.DateTo.Value;
                        if (dtTo.TimeOfDay == TimeSpan.Zero)
                        {
                            dtTo = dtTo.Date.AddDays(1).AddTicks(-1);
                        }
                        string toStr = dtTo.ToString("yyyy-MM-ddTHH:mm:ss.fffffff");
                        sqlParams.Add(new SqliteParameter("@DateTo", toStr));
                    }

                    // Handle term exclusions (-term)
                    for (int i = 0; i < query.ExcludeTerms.Count; i++)
                    {
                        string paramName = $"@ExcludeTerm{i}";
                        whereClauses.Add($"(e.Subject NOT LIKE {paramName} AND e.BodyText NOT LIKE {paramName})");
                        sqlParams.Add(new SqliteParameter(paramName, $"%{query.ExcludeTerms[i]}%"));
                    }

                    string whereSql = whereClauses.Count > 0 ? "WHERE " + string.Join(" AND ", whereClauses) : "";

                    // Total Count Query
                    using (var countCmd = conn.CreateCommand())
                    {
                        countCmd.CommandText = $"SELECT COUNT(*) {fromClause} {whereSql};";
                        foreach (var p in sqlParams) countCmd.Parameters.Add(CloneParameter(p));
                        object? res = countCmd.ExecuteScalar();
                        container.TotalCount = res != null ? Convert.ToInt32(res) : 0;
                    }

                    // Determine Order By clause
                    string orderByClause;
                    if (query.SortOrder == SortOption.OldestFirst)
                    {
                        orderByClause = "ORDER BY e.ReceivedDate ASC";
                    }
                    else if (query.SortOrder == SortOption.NewestFirst)
                    {
                        orderByClause = "ORDER BY e.ReceivedDate DESC";
                    }
                    else
                    {
                        // Default to Relevance
                        if (hasFts)
                        {
                            orderByClause = "ORDER BY fts.rank ASC, e.ReceivedDate DESC";
                        }
                        else
                        {
                            orderByClause = "ORDER BY e.ReceivedDate DESC";
                        }
                    }

                    // Data Retrieval Query
                    var rawResults = new List<SearchResultItem>();
                    int fetchLimit = query.Limit > 0 ? query.Limit : 50;
                    int offset = Math.Max(0, query.Offset);

                    using (var dataCmd = conn.CreateCommand())
                    {
                        dataCmd.CommandText = $@"
                            SELECT e.Id, e.StoreId, e.EntryId, e.Subject, e.SenderName, e.SenderEmail,
                                   e.Recipients, e.ReceivedDate, e.FolderName, e.BodyText,
                                   e.HasAttachment, e.IsRead, e.IsFlagged, e.Importance
                            {fromClause}
                            {whereSql}
                            {orderByClause}
                            LIMIT @Limit OFFSET @Offset;";

                        foreach (var p in sqlParams) dataCmd.Parameters.Add(CloneParameter(p));
                        dataCmd.Parameters.Add(new SqliteParameter("@Limit", fetchLimit));
                        dataCmd.Parameters.Add(new SqliteParameter("@Offset", offset));

                        cancellationToken.ThrowIfCancellationRequested();

                        using (var reader = dataCmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                cancellationToken.ThrowIfCancellationRequested();

                                long emailId = reader.GetInt64(0);
                                string bodyText = reader.IsDBNull(9) ? "" : reader.GetString(9);
                                string snippet = GenerateSnippet(bodyText, query.FreeText);

                                var item = new SearchResultItem
                                {
                                    EmailId = emailId,
                                    StoreId = reader.GetString(1),
                                    EntryId = reader.GetString(2),
                                    Subject = reader.IsDBNull(3) ? "(No Subject)" : reader.GetString(3),
                                    SenderName = reader.IsDBNull(4) ? "" : reader.GetString(4),
                                    SenderEmail = reader.IsDBNull(5) ? "" : reader.GetString(5),
                                    Recipients = reader.IsDBNull(6) ? "" : reader.GetString(6),
                                    ReceivedDate = DateTime.TryParse(reader.GetString(7), out DateTime rdt) ? rdt : DateTime.MinValue,
                                    FolderName = reader.IsDBNull(8) ? "" : reader.GetString(8),
                                    BodySnippet = snippet,
                                    HasAttachment = reader.GetInt32(10) == 1,
                                    IsRead = reader.GetInt32(11) == 1,
                                    IsFlagged = reader.GetInt32(12) == 1,
                                    Importance = reader.GetInt32(13)
                                };

                                if (item.HasAttachment)
                                {
                                    item.AttachmentNames = GetAttachmentNames(conn, emailId);
                                }

                                rawResults.Add(item);
                            }
                        }
                    }

                    container.Items = rawResults;
                    container.Page = query.Page;
                    container.PageSize = fetchLimit;
                    container.HasMoreResults = (offset + rawResults.Count) < container.TotalCount;
                    sw.Stop();
                    container.SearchTimeMs = sw.ElapsedMilliseconds;

                    // Phase 14 — Structured Search Logging
                    Logger.Info($"[SEARCH] Query='{query.RawQuery}'");
                    if (!string.IsNullOrEmpty(query.From)) Logger.Info($"[SEARCH] Parsed From='{query.From}'");
                    if (!string.IsNullOrEmpty(query.To)) Logger.Info($"[SEARCH] Parsed To='{query.To}'");
                    if (!string.IsNullOrEmpty(query.Cc)) Logger.Info($"[SEARCH] Parsed Cc='{query.Cc}'");
                    if (!string.IsNullOrEmpty(query.Subject)) Logger.Info($"[SEARCH] Parsed Subject='{query.Subject}'");
                    if (!string.IsNullOrEmpty(query.Filename)) Logger.Info($"[SEARCH] Parsed Filename='{query.Filename}'");
                    if (!string.IsNullOrEmpty(query.Folder)) Logger.Info($"[SEARCH] Parsed Folder='{query.Folder}'");
                    if (query.HasAttachment.HasValue) Logger.Info($"[SEARCH] Parsed HasAttachment={query.HasAttachment.Value}");
                    if (query.AfterDate.HasValue) Logger.Info($"[SEARCH] Parsed AfterDate={query.AfterDate.Value:yyyy-MM-dd}");
                    if (query.BeforeDate.HasValue) Logger.Info($"[SEARCH] Parsed BeforeDate={query.BeforeDate.Value:yyyy-MM-dd}");
                    Logger.Info($"[SEARCH] Results={container.TotalCount}");
                    Logger.Info($"[SEARCH] Elapsed={sw.ElapsedMilliseconds}ms");
                }

                return container;
            });
        }

        private SqliteParameter CloneParameter(SqliteParameter p)
        {
            return new SqliteParameter(p.ParameterName, p.Value);
        }

        private static bool IsPureNumeric(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return false;
            for (int i = 0; i < token.Length; i++)
            {
                if (!char.IsDigit(token[i]))
                    return false;
            }
            return true;
        }

        private static bool IsEmoji(int codePoint)
        {
            return (codePoint >= 0x1F300 && codePoint <= 0x1FAFF)
                || (codePoint >= 0x2600 && codePoint <= 0x27BF)
                || (codePoint >= 0x1F000 && codePoint <= 0x1F02F)
                || (codePoint >= 0x1F0A0 && codePoint <= 0x1F0FF);
        }

        private static List<string> ExtractEmojis(string text)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(text)) return list;

            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsSurrogatePair(text, i))
                {
                    int codePoint = char.ConvertToUtf32(text, i);
                    if (IsEmoji(codePoint))
                    {
                        list.Add(text.Substring(i, 2));
                    }
                    i++;
                }
                else
                {
                    int codePoint = text[i];
                    if (IsEmoji(codePoint))
                    {
                        list.Add(text.Substring(i, 1));
                    }
                }
            }
            return list;
        }

        private static string RemoveEmojis(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsSurrogatePair(text, i))
                {
                    int codePoint = char.ConvertToUtf32(text, i);
                    if (!IsEmoji(codePoint))
                    {
                        sb.Append(text[i]);
                        sb.Append(text[i + 1]);
                    }
                    i++;
                }
                else
                {
                    int codePoint = text[i];
                    if (!IsEmoji(codePoint))
                    {
                        sb.Append(text[i]);
                    }
                }
            }
            return sb.ToString();
        }

        private List<string> GetKnownFolderNames(SqliteConnection conn)
        {
            lock (_folderLock)
            {
                if (_cachedFolderNames != null && _cachedFolderNames.Count > 0)
                {
                    return _cachedFolderNames;
                }

                var folders = new List<string>();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT DISTINCT FolderName FROM Emails WHERE FolderName IS NOT NULL AND FolderName != '';";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            folders.Add(reader.GetString(0));
                        }
                    }
                }
                _cachedFolderNames = folders;
                return folders;
            }
        }

        private string? BuildFtsCondition(SearchQuery query, List<SqliteParameter> sqlParams)
        {
            var terms = new List<string>();

            if (query.ExactPhrases.Count > 0)
            {
                foreach (var phrase in query.ExactPhrases)
                {
                    string cleanPhrase = phrase.Replace("\"", "").Replace("'", "");
                    cleanPhrase = RemoveEmojis(cleanPhrase).Trim();
                    if (!string.IsNullOrWhiteSpace(cleanPhrase))
                    {
                        terms.Add($"\"{cleanPhrase}\"");
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(query.FreeText))
            {
                string textNoEmoji = RemoveEmojis(query.FreeText);
                string[] parts = textNoEmoji.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in parts)
                {
                    string clean = p.Replace("\"", "").Replace("'", "").Replace("*", "").Trim();
                    if (!string.IsNullOrWhiteSpace(clean) && !clean.Equals("OR", StringComparison.OrdinalIgnoreCase) && !clean.Equals("AND", StringComparison.OrdinalIgnoreCase))
                    {
                        if (IsPureNumeric(clean))
                        {
                            terms.Add($"\"{clean}\"");
                        }
                        else
                        {
                            terms.Add($"\"{clean}\"*");
                        }
                    }
                }
            }

            // Phase 7X: Integrate subject filter into FTS column condition
            if (!string.IsNullOrWhiteSpace(query.SubjectFilter))
            {
                string subjNoEmoji = RemoveEmojis(query.SubjectFilter);
                string[] parts = subjNoEmoji.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                var subjTerms = new List<string>();
                foreach (var p in parts)
                {
                    string clean = p.Replace("\"", "").Replace("'", "").Replace("*", "").Trim();
                    if (!string.IsNullOrWhiteSpace(clean) && !clean.Equals("OR", StringComparison.OrdinalIgnoreCase) && !clean.Equals("AND", StringComparison.OrdinalIgnoreCase))
                    {
                        if (IsPureNumeric(clean))
                        {
                            subjTerms.Add($"\"{clean}\"");
                        }
                        else
                        {
                            subjTerms.Add($"\"{clean}\"*");
                        }
                    }
                }

                if (subjTerms.Count == 1)
                {
                    terms.Add($"Subject: {subjTerms[0]}");
                }
                else if (subjTerms.Count > 1)
                {
                    terms.Add($"Subject: ({string.Join(" ", subjTerms)})");
                }
            }

            if (terms.Count == 0) return null;

            string op = query.BooleanOperator.Equals("OR", StringComparison.OrdinalIgnoreCase) ? " OR " : " AND ";
            string ftsExpr = string.Join(op, terms);

            sqlParams.Add(new SqliteParameter("@FtsQuery", ftsExpr));
            return "EmailSearch MATCH @FtsQuery";
        }

        private string GenerateSnippet(string bodyText, string freeText)
        {
            if (string.IsNullOrWhiteSpace(bodyText)) return string.Empty;

            if (string.IsNullOrWhiteSpace(freeText))
            {
                return bodyText.Length > 160 ? bodyText.Substring(0, 160).Trim() + "..." : bodyText.Trim();
            }

            string lowerBody = bodyText.ToLowerInvariant();
            string lowerQuery = freeText.ToLowerInvariant().Replace("\"", "").Trim();
            string[] terms = lowerQuery.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            int matchIdx = -1;
            foreach (var term in terms)
            {
                if (term.Length > 1)
                {
                    matchIdx = lowerBody.IndexOf(term, StringComparison.Ordinal);
                    if (matchIdx >= 0) break;
                }
            }

            if (matchIdx < 0)
            {
                return bodyText.Length > 160 ? bodyText.Substring(0, 160).Trim() + "..." : bodyText.Trim();
            }

            int start = Math.Max(0, matchIdx - 40);
            int length = Math.Min(bodyText.Length - start, 160);
            string snippet = bodyText.Substring(start, length).Trim();

            if (start > 0) snippet = "..." + snippet;
            if (start + length < bodyText.Length) snippet = snippet + "...";

            return snippet;
        }

        private List<string> GetAttachmentNames(SqliteConnection conn, long emailId)
        {
            var names = new List<string>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Name FROM Attachments WHERE EmailId = @EmailId;";
                cmd.Parameters.AddWithValue("@EmailId", emailId);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        names.Add(reader.GetString(0));
                    }
                }
            }
            return names;
        }

        private List<EmailAttachment> GetEmailAttachments(SqliteConnection conn, long emailId)
        {
            var atts = new List<EmailAttachment>();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Id, EmailId, Name, Extension, Size FROM Attachments WHERE EmailId = @EmailId;";
                cmd.Parameters.AddWithValue("@EmailId", emailId);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        atts.Add(new EmailAttachment
                        {
                            Id = reader.GetInt64(0),
                            EmailId = reader.GetInt64(1),
                            Name = reader.GetString(2),
                            Extension = reader.IsDBNull(3) ? "" : reader.GetString(3),
                            Size = reader.GetInt64(4)
                        });
                    }
                }
            }
            return atts;
        }

        #region Phase 20: Persistent Queue, Failed Item Tracking & Self-Healing

        public Task EnqueuePendingChangeAsync(string storeId, string entryId, string operation, string? folderEntryId)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        string nowStr = DateTime.Now.ToString("o");
                        long existingId = 0;
                        using (var check = conn.CreateCommand())
                        {
                            check.Transaction = tx;
                            check.CommandText = "SELECT Id FROM PendingChanges WHERE StoreId = @StoreId AND EntryId = @EntryId;";
                            check.Parameters.AddWithValue("@StoreId", storeId);
                            check.Parameters.AddWithValue("@EntryId", entryId);
                            object? res = check.ExecuteScalar();
                            if (res != null && res != DBNull.Value) existingId = Convert.ToInt64(res);
                        }

                        if (existingId > 0)
                        {
                            using (var update = conn.CreateCommand())
                            {
                                update.Transaction = tx;
                                update.CommandText = @"
                                    UPDATE PendingChanges SET
                                        Operation = @Operation,
                                        FolderEntryId = @FolderEntryId,
                                        UpdatedAt = @UpdatedAt,
                                        Status = CASE WHEN Status IN ('Failed', 'Completed', 'Processing') THEN 'Pending' ELSE Status END,
                                        AttemptCount = CASE WHEN Status IN ('Failed', 'Completed') THEN 0 ELSE AttemptCount END,
                                        LastError = CASE WHEN Status IN ('Failed', 'Completed') THEN NULL ELSE LastError END
                                    WHERE Id = @Id;";
                                update.Parameters.AddWithValue("@Operation", operation);
                                update.Parameters.AddWithValue("@FolderEntryId", (object?)folderEntryId ?? DBNull.Value);
                                update.Parameters.AddWithValue("@UpdatedAt", nowStr);
                                update.Parameters.AddWithValue("@Id", existingId);
                                update.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            using (var insert = conn.CreateCommand())
                            {
                                insert.Transaction = tx;
                                insert.CommandText = @"
                                    INSERT INTO PendingChanges (
                                        StoreId, EntryId, Operation, FolderEntryId, CreatedAt, UpdatedAt, AttemptCount, Status, LastError
                                    ) VALUES (
                                        @StoreId, @EntryId, @Operation, @FolderEntryId, @CreatedAt, @UpdatedAt, 0, 'Pending', NULL
                                    );";
                                insert.Parameters.AddWithValue("@StoreId", storeId);
                                insert.Parameters.AddWithValue("@EntryId", entryId);
                                insert.Parameters.AddWithValue("@Operation", operation);
                                insert.Parameters.AddWithValue("@FolderEntryId", (object?)folderEntryId ?? DBNull.Value);
                                insert.Parameters.AddWithValue("@CreatedAt", nowStr);
                                insert.Parameters.AddWithValue("@UpdatedAt", nowStr);
                                insert.ExecuteNonQuery();
                            }
                        }
                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error($"[QUEUE] Failed to enqueue pending change for StoreId={storeId} EntryId={entryId}", ex);
                    }
                }
            });
        }

        public Task<List<PendingChangeItem>> GetPendingChangesAsync(int maxCount = 50)
        {
            return Task.Run(() =>
            {
                var list = new List<PendingChangeItem>();
                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            cmd.CommandText = @"
                                SELECT Id, StoreId, EntryId, Operation, FolderEntryId, CreatedAt, UpdatedAt, AttemptCount, Status, LastError
                                FROM PendingChanges
                                WHERE Status = 'Pending'
                                ORDER BY Id ASC
                                LIMIT @Limit;";
                            cmd.Parameters.AddWithValue("@Limit", maxCount);

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    list.Add(new PendingChangeItem
                                    {
                                        Id = reader.GetInt64(0),
                                        StoreId = reader.GetString(1),
                                        EntryId = reader.GetString(2),
                                        Operation = reader.GetString(3),
                                        FolderEntryId = reader.IsDBNull(4) ? null : reader.GetString(4),
                                        CreatedAt = DateTime.TryParse(reader.GetString(5), out DateTime cat) ? cat : DateTime.Now,
                                        UpdatedAt = DateTime.TryParse(reader.GetString(6), out DateTime uat) ? uat : DateTime.Now,
                                        AttemptCount = reader.GetInt32(7),
                                        Status = "Processing",
                                        LastError = reader.IsDBNull(9) ? null : reader.GetString(9)
                                    });
                                }
                            }
                        }

                        if (list.Count > 0)
                        {
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tx;
                                string ids = string.Join(",", list.Select(x => x.Id));
                                cmd.CommandText = $"UPDATE PendingChanges SET Status = 'Processing', UpdatedAt = @UpdatedAt WHERE Id IN ({ids});";
                                cmd.Parameters.AddWithValue("@UpdatedAt", DateTime.Now.ToString("o"));
                                cmd.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error("Failed to fetch pending changes", ex);
                    }
                }
                return list;
            });
        }

        public Task<int> GetPendingChangesCountAsync()
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM PendingChanges WHERE Status IN ('Pending', 'Processing');";
                    var result = cmd.ExecuteScalar();
                    return Convert.ToInt32(result);
                }
            });
        }

        public Task<int> IncrementAttemptCountAsync(long id)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE PendingChanges SET AttemptCount = AttemptCount + 1, UpdatedAt = @UpdatedAt WHERE Id = @Id;";
                    cmd.Parameters.AddWithValue("@UpdatedAt", DateTime.Now.ToString("o"));
                    cmd.Parameters.AddWithValue("@Id", id);
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "SELECT AttemptCount FROM PendingChanges WHERE Id = @Id;";
                    object? res = cmd.ExecuteScalar();
                    return res != null ? Convert.ToInt32(res) : 0;
                }
            });
        }

        public Task MarkPendingChangeStatusAsync(long id, string status, string? error = null)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE PendingChanges SET
                            Status = @Status,
                            LastError = @LastError,
                            UpdatedAt = @UpdatedAt
                        WHERE Id = @Id AND Status = 'Processing';";
                    cmd.Parameters.AddWithValue("@Status", status);
                    cmd.Parameters.AddWithValue("@LastError", (object?)error ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@UpdatedAt", DateTime.Now.ToString("o"));
                    cmd.Parameters.AddWithValue("@Id", id);
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task MarkPendingChangesCompletedBatchAsync(IEnumerable<long> ids)
        {
            return Task.Run(() =>
            {
                var idList = ids.ToList();
                if (!idList.Any()) return;

                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        string nowStr = DateTime.Now.ToString("o");
                        foreach (var id in idList)
                        {
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tx;
                                cmd.CommandText = "UPDATE PendingChanges SET Status = 'Completed', UpdatedAt = @UpdatedAt WHERE Id = @Id AND Status = 'Processing';";
                                cmd.Parameters.AddWithValue("@UpdatedAt", nowStr);
                                cmd.Parameters.AddWithValue("@Id", id);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error("Error updating batch completed pending changes", ex);
                    }
                }
            });
        }

        public Task RecoverStaleProcessingChangesAsync()
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    var threshold = DateTime.Now.AddMinutes(-5).ToString("o");
                    cmd.CommandText = "UPDATE PendingChanges SET Status = 'Pending', UpdatedAt = @UpdatedAt WHERE Status = 'Processing' AND UpdatedAt < @Threshold;";
                    cmd.Parameters.AddWithValue("@UpdatedAt", DateTime.Now.ToString("o"));
                    cmd.Parameters.AddWithValue("@Threshold", threshold);
                    int count = cmd.ExecuteNonQuery();
                    if (count > 0)
                    {
                        Logger.Info($"[QUEUE] Recovered {count} stale processing operations back to Pending.");
                    }
                }
            });
        }

        public Task SaveFailedIndexItemAsync(string storeId, string entryId, string operation, int attemptCount, string lastError)
        {
            return SaveFailedIndexItemAsync(new FailedIndexItem
            {
                StoreId = storeId,
                EntryId = entryId,
                OperationType = operation,
                AttemptCount = attemptCount,
                ErrorMessage = lastError ?? string.Empty,
                FailureReason = lastError ?? string.Empty
            });
        }

        public Task SaveFailedIndexItemAsync(FailedIndexItem item)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    string nowStr = DateTime.Now.ToString("o");
                    cmd.CommandText = @"
                        INSERT INTO FailedIndexItems (
                            StoreId, EntryId, FolderName, OperationType, FailureReason, ExceptionType, AttemptCount, FirstFailedAt, LastFailedAt, LastRetryAt, Status, ErrorMessage, Resolved
                        ) VALUES (
                            @StoreId, @EntryId, @FolderName, @OperationType, @FailureReason, @ExceptionType, @AttemptCount, @FirstFailedAt, @LastFailedAt, @LastRetryAt, @Status, @ErrorMessage, 0
                        )
                        ON CONFLICT(StoreId, EntryId) DO UPDATE SET
                            FolderName = excluded.FolderName,
                            OperationType = excluded.OperationType,
                            FailureReason = excluded.FailureReason,
                            ExceptionType = excluded.ExceptionType,
                            AttemptCount = FailedIndexItems.AttemptCount + 1,
                            LastFailedAt = excluded.LastFailedAt,
                            LastRetryAt = excluded.LastRetryAt,
                            Status = excluded.Status,
                            ErrorMessage = excluded.ErrorMessage,
                            Resolved = 0;";
                    cmd.Parameters.AddWithValue("@StoreId", item.StoreId);
                    cmd.Parameters.AddWithValue("@EntryId", item.EntryId);
                    cmd.Parameters.AddWithValue("@FolderName", (object?)item.FolderName ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@OperationType", item.OperationType ?? "Add");
                    cmd.Parameters.AddWithValue("@FailureReason", (object?)item.FailureReason ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ExceptionType", (object?)item.ExceptionType ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@AttemptCount", item.AttemptCount);
                    cmd.Parameters.AddWithValue("@FirstFailedAt", item.FirstFailedAt.ToString("o"));
                    cmd.Parameters.AddWithValue("@LastFailedAt", nowStr);
                    cmd.Parameters.AddWithValue("@LastRetryAt", item.LastRetryAt.HasValue ? item.LastRetryAt.Value.ToString("o") : DBNull.Value);
                    cmd.Parameters.AddWithValue("@Status", item.Status ?? "Failed");
                    cmd.Parameters.AddWithValue("@ErrorMessage", (object?)item.ErrorMessage ?? DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task<List<FailedIndexItem>> GetFailedIndexItemsAsync(bool includeResolved = false)
        {
            return GetFailedIndexItemsAsync(includeResolved ? null : "Failed", 500);
        }

        public Task<List<FailedIndexItem>> GetFailedIndexItemsAsync(string? status, int maxCount = 500)
        {
            return Task.Run(() =>
            {
                var list = new List<FailedIndexItem>();
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    if (string.IsNullOrEmpty(status))
                    {
                        cmd.CommandText = "SELECT Id, StoreId, EntryId, FolderName, OperationType, FailureReason, ExceptionType, AttemptCount, FirstFailedAt, LastFailedAt, LastRetryAt, Status, ErrorMessage, Resolved FROM FailedIndexItems ORDER BY Id DESC LIMIT @Limit;";
                    }
                    else
                    {
                        cmd.CommandText = "SELECT Id, StoreId, EntryId, FolderName, OperationType, FailureReason, ExceptionType, AttemptCount, FirstFailedAt, LastFailedAt, LastRetryAt, Status, ErrorMessage, Resolved FROM FailedIndexItems WHERE Status = @Status ORDER BY Id DESC LIMIT @Limit;";
                        cmd.Parameters.AddWithValue("@Status", status);
                    }
                    cmd.Parameters.AddWithValue("@Limit", maxCount);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(new FailedIndexItem
                            {
                                Id = reader.GetInt64(0),
                                StoreId = reader.GetString(1),
                                EntryId = reader.GetString(2),
                                FolderName = reader.IsDBNull(3) ? "" : reader.GetString(3),
                                OperationType = reader.IsDBNull(4) ? "Add" : reader.GetString(4),
                                FailureReason = reader.IsDBNull(5) ? "" : reader.GetString(5),
                                ExceptionType = reader.IsDBNull(6) ? "" : reader.GetString(6),
                                AttemptCount = reader.GetInt32(7),
                                FirstFailedAt = DateTime.TryParse(reader.GetString(8), out DateTime fdt) ? fdt : DateTime.Now,
                                LastFailedAt = DateTime.TryParse(reader.GetString(9), out DateTime ldt) ? ldt : DateTime.Now,
                                LastRetryAt = !reader.IsDBNull(10) && DateTime.TryParse(reader.GetString(10), out DateTime rdt) ? rdt : null,
                                Status = reader.IsDBNull(11) ? "Failed" : reader.GetString(11),
                                ErrorMessage = reader.IsDBNull(12) ? "" : reader.GetString(12),
                                Resolved = reader.GetInt32(13) == 1
                            });
                        }
                    }
                }
                return list;
            });
        }

        public Task<int> GetFailedIndexItemsCountAsync(string? status = "Failed")
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    if (string.IsNullOrEmpty(status))
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM FailedIndexItems;";
                    }
                    else
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM FailedIndexItems WHERE Status = @Status;";
                        cmd.Parameters.AddWithValue("@Status", status);
                    }
                    var result = cmd.ExecuteScalar();
                    return Convert.ToInt32(result);
                }
            });
        }

        public Task UpdateFailedIndexItemStatusAsync(long id, string status, string? error = null)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE FailedIndexItems SET
                            Status = @Status,
                            ErrorMessage = COALESCE(@Error, ErrorMessage),
                            LastRetryAt = @Now,
                            Resolved = CASE WHEN @Status = 'Recovered' OR @Status = 'Ignored' THEN 1 ELSE 0 END
                        WHERE Id = @Id;";
                    cmd.Parameters.AddWithValue("@Status", status);
                    cmd.Parameters.AddWithValue("@Error", (object?)error ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Now", DateTime.Now.ToString("o"));
                    cmd.Parameters.AddWithValue("@Id", id);
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task RemoveFailedIndexItemAsync(long id)
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "DELETE FROM FailedIndexItems WHERE Id = @Id;";
                    cmd.Parameters.AddWithValue("@Id", id);
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task RemoveAllFailedIndexItemsAsync()
        {
            return Task.Run(() =>
            {
                using (var conn = _factory.CreateConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "DELETE FROM FailedIndexItems;";
                    cmd.ExecuteNonQuery();
                }
            });
        }

        public Task<double> GetDatabaseSizeMbAsync()
        {
            return Task.Run(() =>
            {
                try
                {
                    string dbPath = _factory.DbPath;
                    if (File.Exists(dbPath))
                    {
                        long bytes = new FileInfo(dbPath).Length;
                        return bytes / (1024.0 * 1024.0);
                    }
                }
                catch { }
                return 0.0;
            });
        }

        public Task<IndexIntegrityResult> ScanIndexIntegrityAsync(CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                var result = new IndexIntegrityResult();
                Logger.Info("[INTEGRITY] Scan started.");

                using (var conn = _factory.CreateConnection())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "PRAGMA integrity_check;";
                        object? res = cmd.ExecuteScalar();
                        string status = res?.ToString() ?? "ok";
                        if (!string.Equals(status, "ok", StringComparison.OrdinalIgnoreCase))
                        {
                            result.InvalidRecordCount += 1;
                            Logger.Warn($"[INTEGRITY] PRAGMA integrity_check failed: {status}");
                        }
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM Emails;";
                        object? res = cmd.ExecuteScalar();
                        result.EmailCount = res != null ? Convert.ToInt32(res) : 0;
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM EmailSearch;";
                        object? res = cmd.ExecuteScalar();
                        result.FtsCount = res != null ? Convert.ToInt32(res) : 0;
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM Emails WHERE CAST(Id AS TEXT) NOT IN (SELECT EmailId FROM EmailSearch);";
                        object? res = cmd.ExecuteScalar();
                        result.MissingFtsCount = res != null ? Convert.ToInt32(res) : 0;
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM EmailSearch WHERE EmailId NOT IN (SELECT CAST(Id AS TEXT) FROM Emails);";
                        object? res = cmd.ExecuteScalar();
                        result.OrphanFtsCount = res != null ? Convert.ToInt32(res) : 0;
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM (SELECT StoreId, EntryId FROM Emails GROUP BY StoreId, EntryId HAVING COUNT(*) > 1);";
                        object? res = cmd.ExecuteScalar();
                        result.DuplicateCount = res != null ? Convert.ToInt32(res) : 0;
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM Emails WHERE StoreId IS NULL OR StoreId = '' OR EntryId IS NULL OR EntryId = '';";
                        object? res = cmd.ExecuteScalar();
                        result.InvalidRecordCount = res != null ? Convert.ToInt32(res) : 0;
                    }
                }

                Logger.Info($"[INTEGRITY] Missing FTS records detected: {result.MissingFtsCount}");
                Logger.Info($"[INTEGRITY] Orphan FTS records detected: {result.OrphanFtsCount}");
                Logger.Info($"[INTEGRITY] Duplicate records detected: {result.DuplicateCount}");
                return result;
            }, cancellationToken);
        }

        public Task RepairIndexIntegrityAsync(IndexIntegrityResult scanResult, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                if (scanResult == null || scanResult.IsHealthy) return;

                using (var conn = _factory.CreateConnection())
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // 1. Repair missing FTS entries
                        if (scanResult.MissingFtsCount > 0)
                        {
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tx;
                                cmd.CommandText = @"
                                    INSERT INTO EmailSearch (EmailId, Subject, SenderName, SenderEmail, Recipients, BodyText, AttachmentNames)
                                    SELECT CAST(e.Id AS TEXT), COALESCE(e.Subject, ''), COALESCE(e.SenderName, ''), COALESCE(e.SenderEmail, ''), COALESCE(e.Recipients, ''), COALESCE(e.BodyText, ''), ''
                                    FROM Emails e
                                    WHERE CAST(e.Id AS TEXT) NOT IN (SELECT EmailId FROM EmailSearch);";
                                cmd.ExecuteNonQuery();
                            }
                        }

                        cancellationToken.ThrowIfCancellationRequested();

                        // 2. Remove orphan FTS entries
                        if (scanResult.OrphanFtsCount > 0)
                        {
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tx;
                                cmd.CommandText = "DELETE FROM EmailSearch WHERE EmailId NOT IN (SELECT CAST(Id AS TEXT) FROM Emails);";
                                cmd.ExecuteNonQuery();
                            }
                        }

                        cancellationToken.ThrowIfCancellationRequested();

                        // 3. Remove duplicate logical emails (keeping highest Id)
                        if (scanResult.DuplicateCount > 0)
                        {
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tx;
                                cmd.CommandText = @"
                                    DELETE FROM Emails
                                    WHERE Id NOT IN (
                                        SELECT MAX(Id) FROM Emails GROUP BY StoreId, EntryId
                                    );
                                    DELETE FROM EmailSearch
                                    WHERE EmailId NOT IN (
                                        SELECT CAST(Id AS TEXT) FROM Emails
                                    );";
                                cmd.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                        lock (_folderLock) { _cachedFolderNames = null; }
                        Logger.Info("[INTEGRITY] Repair completed successfully.");
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        Logger.Error("[INTEGRITY ERROR] Failed to perform self-healing repair transaction", ex);
                        throw;
                    }
                }
            }, cancellationToken);
        }

        #endregion

        private EmailDocument ReadEmailDocument(SqliteDataReader reader)
        {
            return new EmailDocument
            {
                Id = reader.GetInt64(reader.GetOrdinal("Id")),
                StoreId = reader.GetString(reader.GetOrdinal("StoreId")),
                EntryId = reader.GetString(reader.GetOrdinal("EntryId")),
                ConversationId = reader.IsDBNull(reader.GetOrdinal("ConversationId")) ? "" : reader.GetString(reader.GetOrdinal("ConversationId")),
                Subject = reader.IsDBNull(reader.GetOrdinal("Subject")) ? "" : reader.GetString(reader.GetOrdinal("Subject")),
                SenderName = reader.IsDBNull(reader.GetOrdinal("SenderName")) ? "" : reader.GetString(reader.GetOrdinal("SenderName")),
                SenderEmail = reader.IsDBNull(reader.GetOrdinal("SenderEmail")) ? "" : reader.GetString(reader.GetOrdinal("SenderEmail")),
                Recipients = reader.IsDBNull(reader.GetOrdinal("Recipients")) ? "" : reader.GetString(reader.GetOrdinal("Recipients")),
                Cc = reader.IsDBNull(reader.GetOrdinal("Cc")) ? "" : reader.GetString(reader.GetOrdinal("Cc")),
                Bcc = reader.IsDBNull(reader.GetOrdinal("Bcc")) ? "" : reader.GetString(reader.GetOrdinal("Bcc")),
                BodyText = reader.IsDBNull(reader.GetOrdinal("BodyText")) ? "" : reader.GetString(reader.GetOrdinal("BodyText")),
                ReceivedDate = DateTime.TryParse(reader.GetString(reader.GetOrdinal("ReceivedDate")), out DateTime rdt) ? rdt : DateTime.MinValue,
                SentDate = DateTime.TryParse(reader.GetString(reader.GetOrdinal("SentDate")), out DateTime sdt) ? sdt : DateTime.MinValue,
                FolderName = reader.IsDBNull(reader.GetOrdinal("FolderName")) ? "" : reader.GetString(reader.GetOrdinal("FolderName")),
                FolderEntryId = reader.IsDBNull(reader.GetOrdinal("FolderEntryId")) ? "" : reader.GetString(reader.GetOrdinal("FolderEntryId")),
                HasAttachment = reader.GetInt32(reader.GetOrdinal("HasAttachment")) == 1,
                IsRead = reader.GetInt32(reader.GetOrdinal("IsRead")) == 1,
                IsFlagged = reader.GetInt32(reader.GetOrdinal("IsFlagged")) == 1,
                Importance = reader.GetInt32(reader.GetOrdinal("Importance")),
                Categories = reader.IsDBNull(reader.GetOrdinal("Categories")) ? "" : reader.GetString(reader.GetOrdinal("Categories")),
                Size = reader.GetInt64(reader.GetOrdinal("Size")),
                LastModified = DateTime.TryParse(reader.GetString(reader.GetOrdinal("LastModified")), out DateTime mdt) ? mdt : DateTime.MinValue,
                IndexedDate = DateTime.TryParse(reader.GetString(reader.GetOrdinal("IndexedDate")), out DateTime idt) ? idt : DateTime.MinValue,
            };
        }
    }
}
