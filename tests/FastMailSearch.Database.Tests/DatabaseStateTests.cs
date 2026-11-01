using System;
using System.IO;
using System.Threading.Tasks;
using FastMailSearch.Infrastructure.Database;
using Xunit;

namespace FastMailSearch.Database.Tests
{
    public class DatabaseStateTests : IDisposable
    {
        private readonly string _dbPath;
        private readonly DatabaseConnectionFactory _factory;
        private readonly SqliteEmailRepository _repository;

        public DatabaseStateTests()
        {
            _dbPath = Path.Combine(Path.GetTempPath(), $"DbStateTest_{Guid.NewGuid():N}.db");
            _factory = new DatabaseConnectionFactory(_dbPath);
            _repository = new SqliteEmailRepository(_factory);
            _repository.InitializeDatabase();
        }

        public void Dispose()
        {
            try { if (File.Exists(_dbPath)) File.Delete(_dbPath); } catch { }
        }

        [Fact]
        public async Task SaveIndexStateAsync_PersistsToDatabase()
        {
            await _repository.SaveIndexStateAsync("Indexing", 500, 1000, "Inbox", null);
            var state = await _repository.GetPersistedIndexStateAsync();

            Assert.True(state.ContainsKey("IndexState"));
            Assert.Equal("Indexing", state["IndexState"]);
            Assert.Equal("500", state["LastProcessedCount"]);
            Assert.Equal("1000", state["LastTotalCount"]);
            Assert.Equal("Inbox", state["LastIndexedFolder"]);
        }

        [Fact]
        public async Task SaveIndexStateAsync_UpdatesOnCompletion()
        {
            await _repository.SaveIndexStateAsync("Indexing", 100, 1000, "Inbox", null);
            await _repository.SaveIndexStateAsync("Completed", 1000, 1000, "Sent", null);

            var state = await _repository.GetPersistedIndexStateAsync();

            Assert.Equal("Completed", state["IndexState"]);
            Assert.Equal("1000", state["LastProcessedCount"]);
            Assert.True(state.ContainsKey("IndexCompletedAt"));
        }
    }
}
