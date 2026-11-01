using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Infrastructure.Indexing;
using Microsoft.Data.Sqlite;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class Phase29RegressionTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public Phase29RegressionTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"Phase29Tests_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
        }

        public void Dispose()
        {
            try { if (File.Exists(_dbPath)) File.Delete(_dbPath); } catch { }
        }

        [Fact]
        public async Task Test01_Migration_ResolvesDuplicatesSafely()
        {
            // Create legacy DB without unique constraint
            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
                    CREATE TABLE FailedIndexItems (
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
                        Resolved INTEGER NOT NULL DEFAULT 0
                    );
                    INSERT INTO FailedIndexItems (StoreId, EntryId, OperationType, AttemptCount, FirstFailedAt, LastFailedAt) 
                    VALUES ('A', 'B', 'Add', 1, '2020-01-01T00:00:00.0000000', '2020-01-02T00:00:00.0000000');
                    INSERT INTO FailedIndexItems (StoreId, EntryId, OperationType, AttemptCount, FirstFailedAt, LastFailedAt) 
                    VALUES ('A', 'B', 'Add', 2, '2020-01-03T00:00:00.0000000', '2020-01-04T00:00:00.0000000');
                ";
                cmd.ExecuteNonQuery();
            }

            var initializer = new DatabaseInitializer(_factory);
            initializer.Initialize(); // This should trigger the migration

            var failedItems = await _repository.GetFailedIndexItemsAsync("Failed", 100);
            Assert.Single(failedItems); // Only 1 left
            Assert.Equal(3, failedItems[0].AttemptCount); // 1 + 2
            Assert.Equal("2020-01-01", failedItems[0].FirstFailedAt.ToString("yyyy-MM-dd"));
        }

        [Fact]
        public async Task Test02_Migration_IsIdempotent()
        {
            var initializer = new DatabaseInitializer(_factory);
            initializer.Initialize();
            initializer.Initialize(); // Should not fail

            await _repository.SaveFailedIndexItemAsync("Store", "Entry", "Add", 1, "Error");
            await _repository.SaveFailedIndexItemAsync("Store", "Entry", "Add", 1, "Error");

            var count = await _repository.GetFailedIndexItemsCountAsync("Failed");
            Assert.Equal(1, count); // Upsert works
        }

        [Fact]
        public async Task Test03_QueueRaceCondition_ProcessingDoesNotOverwrite()
        {
            var initializer = new DatabaseInitializer(_factory);
            initializer.Initialize();

            await _repository.EnqueuePendingChangeAsync("S1", "E1", "Add", null);
            var pending = await _repository.GetPendingChangesAsync(1); // Changes it to 'Processing' in DB
            Assert.Single(pending);
            Assert.Equal("Add", pending[0].Operation);

            // While processing, user moves the item:
            await _repository.EnqueuePendingChangeAsync("S1", "E1", "Move", "Folder2");

            // Indexer finishes the first 'Add' event
            await _repository.MarkPendingChangesCompletedBatchAsync(new[] { pending[0].Id });

            // Check if DB retained the new 'Move' event as Pending
            using (var conn = _factory.CreateConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Operation, Status FROM PendingChanges WHERE Id = @Id";
                cmd.Parameters.AddWithValue("@Id", pending[0].Id);
                using (var reader = cmd.ExecuteReader())
                {
                    Assert.True(reader.Read());
                    Assert.Equal("Move", reader.GetString(0));
                    Assert.Equal("Pending", reader.GetString(1));
                }
            }
        }
    }
}
