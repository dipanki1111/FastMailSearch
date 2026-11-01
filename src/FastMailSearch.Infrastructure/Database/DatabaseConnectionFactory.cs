using System;
using System.IO;
using FastMailSearch.Core.Models;
using Microsoft.Data.Sqlite;
using SQLitePCL;

namespace FastMailSearch.Infrastructure.Database
{
    public class DatabaseConnectionFactory
    {
        private readonly string _dbPath;

        static DatabaseConnectionFactory()
        {
            Batteries.Init();
        }

        public DatabaseConnectionFactory(string dbPath)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                dbPath = AppSettings.GetDatabasePath();
            }
            else
            {
                AppSettings.ValidateDatabasePath(dbPath);
            }

            _dbPath = dbPath;

            string? dir = Path.GetDirectoryName(_dbPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            try
            {
                bool exists = File.Exists(_dbPath);
                long bytes = exists ? new FileInfo(_dbPath).Length : 0;
                double sizeMb = bytes / (1024.0 * 1024.0);
                Logging.Logger.Info($"[DATABASE PATH] Selected DB: '{_dbPath}' (Exists={exists}, Size={sizeMb:F2} MB, Bytes={bytes})");

                if (!exists)
                {
                    Logging.Logger.Warn($"[DATABASE PATH] Database does not exist at '{_dbPath}'. An empty database will be created upon first connection if not an intended path.");
                }
            }
            catch (Exception ex)
            {
                Logging.Logger.Warn($"[DATABASE PATH] Error inspecting DB path '{_dbPath}': {ex.Message}");
            }
        }

        public string DbPath => _dbPath;

        public SqliteConnection CreateConnection()
        {
            AppSettings.ValidateDatabasePath(_dbPath);

            var builder = new SqliteConnectionStringBuilder
            {
                DataSource = _dbPath,
                Mode = SqliteOpenMode.ReadWriteCreate,
                Cache = SqliteCacheMode.Shared,
                DefaultTimeout = 30
            };

            var connection = new SqliteConnection(builder.ConnectionString);
            connection.Open();

            using (var cmd = connection.CreateCommand())
            {
                cmd.CommandText = "PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL; PRAGMA foreign_keys=ON; PRAGMA temp_store=MEMORY; PRAGMA busy_timeout=30000;";
                cmd.ExecuteNonQuery();
            }

            return connection;
        }
    }
}
