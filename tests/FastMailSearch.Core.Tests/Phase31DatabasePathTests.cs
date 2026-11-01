using System;
using System.IO;
using Xunit;
using FastMailSearch.Infrastructure.Database;
using FastMailSearch.Core.Models;

namespace FastMailSearch.Core.Tests
{
    [Collection("Non-Parallel Environment Tests")]
    public class Phase31DatabasePathTests
    {
        private static string GetExpectedDefaultPath()
        {
            string docsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(docsFolder, "FastMailSearch", "FastMailSearch.db");
        }

        [Fact]
        public void AppSettings_GetDatabasePath_DefaultsToMyDocuments_WhenEnvVarUnset()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, null);

                string resolved = AppSettings.GetDatabasePath();

                Assert.Equal(GetExpectedDefaultPath(), resolved);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("\t\r\n")]
        public void AppSettings_GetDatabasePath_DefaultsToMyDocuments_WhenEnvVarEmptyOrWhitespace(string blankValue)
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, blankValue);

                string resolved = AppSettings.GetDatabasePath();

                Assert.Equal(GetExpectedDefaultPath(), resolved);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Fact]
        public void AppSettings_GetDatabasePath_HonorsOverride_WhenEnvVarSet()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                string overridePath = Path.Combine(Path.GetTempPath(), "FmsTestIsolated.db");
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, overridePath);

                string resolved = AppSettings.GetDatabasePath();

                Assert.Equal(overridePath, resolved);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Fact]
        public void AppSettings_GetDatabasePath_TrimsWhitespace_WhenEnvVarHasSurroundingWhitespace()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                string expectedPath = Path.Combine(Path.GetTempPath(), "FmsTrimTest.db");
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, $"  {expectedPath}  \t");

                string resolved = AppSettings.GetDatabasePath();

                Assert.Equal(expectedPath, resolved);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Fact]
        public void AppSettings_DatabasePath_DefaultsToMyDocuments()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, null);

                var settings = new AppSettings();

                Assert.Equal(GetExpectedDefaultPath(), settings.DatabasePath);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Fact]
        public void AppSettings_DatabasePath_HonorsOverride_WhenEnvVarSet()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                string overridePath = Path.Combine(Path.GetTempPath(), "FmsSettingsOverride.db");
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, overridePath);

                var settings = new AppSettings();

                Assert.Equal(overridePath, settings.DatabasePath);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Fact]
        public void DatabaseConnectionFactory_NullPath_DefaultsToMyDocuments()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, null);

                var factory = new DatabaseConnectionFactory(null);

                Assert.Equal(GetExpectedDefaultPath(), factory.DbPath);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        public void DatabaseConnectionFactory_BlankPath_DefaultsToMyDocuments(string blankPath)
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, null);

                var factory = new DatabaseConnectionFactory(blankPath);

                Assert.Equal(GetExpectedDefaultPath(), factory.DbPath);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Fact]
        public void DatabaseConnectionFactory_NullPath_HonorsOverride_WhenEnvVarSet()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                string overridePath = Path.Combine(Path.GetTempPath(), "FmsFactoryOverride.db");
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, overridePath);

                var factory = new DatabaseConnectionFactory(null);

                Assert.Equal(overridePath, factory.DbPath);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Fact]
        public void DatabaseConnectionFactory_ExplicitPath_HonorsExplicitPath_EvenWhenEnvVarSet()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                string overridePath = Path.Combine(Path.GetTempPath(), "FmsEnvOverride.db");
                string explicitPath = Path.Combine(Path.GetTempPath(), "FmsExplicitPath.db");
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, overridePath);

                var factory = new DatabaseConnectionFactory(explicitPath);

                Assert.Equal(explicitPath, factory.DbPath);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }

        [Fact]
        public void SqliteEmailRepository_DatabasePath_ReflectsFactoryResolvedPath()
        {
            string originalEnv = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
            try
            {
                string overridePath = Path.Combine(Path.GetTempPath(), "FmsRepoOverride.db");
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, overridePath);

                var factory = new DatabaseConnectionFactory(null);
                var repo = new SqliteEmailRepository(factory);

                Assert.Equal(overridePath, repo.DatabasePath);
            }
            finally
            {
                Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, originalEnv);
            }
        }
    }
}
