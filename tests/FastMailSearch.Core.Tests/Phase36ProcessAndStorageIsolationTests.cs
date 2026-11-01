using System;
using System.Diagnostics;
using System.IO;
using Xunit;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Logging;
using FastMailSearch.OutlookAddIn;
using FastMailSearch.UI;

namespace FastMailSearch.Core.Tests
{
    [Collection("Non-Parallel Environment Tests")]
    public class Phase36ProcessAndStorageIsolationTests
    {
        private static readonly object _envLock = new object();

        [Fact]
        public void Logger_GetLogDirectory_DefaultsToLocalAppData_WhenEnvVarUnset()
        {
            lock (_envLock)
            {
                string? origEnv = Environment.GetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable);
                try
                {
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, null);
                    Logger.ResetLogDirectory();

                    string resolved = Logger.GetLogDirectory();
                    string expected = Logger.GetDefaultLogDirectory();

                    Assert.Equal(expected, resolved);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, origEnv);
                    Logger.ResetLogDirectory();
                }
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("\t\r\n")]
        public void Logger_GetLogDirectory_DefaultsToLocalAppData_WhenEnvVarWhitespace(string blankVal)
        {
            lock (_envLock)
            {
                string? origEnv = Environment.GetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable);
                try
                {
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, blankVal);
                    Logger.ResetLogDirectory();

                    string resolved = Logger.GetLogDirectory();
                    string expected = Logger.GetDefaultLogDirectory();

                    Assert.Equal(expected, resolved);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, origEnv);
                    Logger.ResetLogDirectory();
                }
            }
        }

        [Fact]
        public void Logger_GetLogDirectory_HonorsOverride_WhenEnvVarSet()
        {
            lock (_envLock)
            {
                string? origEnv = Environment.GetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable);
                try
                {
                    string customLogDir = Path.Combine(Path.GetTempPath(), "FmsCustomLogs");
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, customLogDir);
                    Logger.ResetLogDirectory();

                    string resolved = Logger.GetLogDirectory();

                    Assert.Equal(customLogDir, resolved);
                    Assert.Equal(customLogDir, Logger.LogDirectory);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, origEnv);
                    Logger.ResetLogDirectory();
                }
            }
        }

        [Fact]
        public void Logger_Initialize_WithExplicitCustomDir_OverridesEnvironmentAndDefault()
        {
            lock (_envLock)
            {
                string? origEnv = Environment.GetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable);
                try
                {
                    string envDir = Path.Combine(Path.GetTempPath(), "FmsEnvLogs");
                    string explicitDir = Path.Combine(Path.GetTempPath(), "FmsExplicitLogs");
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, envDir);

                    Logger.Initialize(explicitDir);

                    Assert.Equal(explicitDir, Logger.LogDirectory);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, origEnv);
                    Logger.ResetLogDirectory();
                }
            }
        }

        [Fact]
        public void AppSettings_NormalizePath_HandlesCaseTrailingSeparatorsAndRelativeDots()
        {
            string raw = @"C:\Users\TestUser\Documents\..\Documents\FastMailSearch\FastMailSearch.db\";
            string normalized = AppSettings.NormalizePath(raw);

            string expected = Path.GetFullPath(@"C:\Users\TestUser\Documents\FastMailSearch\FastMailSearch.db");
            Assert.Equal(expected, normalized, ignoreCase: true);
        }

        [Fact]
        public void AppSettings_RequireOverride_ThrowsException_WhenProductionDatabaseResolved()
        {
            lock (_envLock)
            {
                string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
                string? origDb = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
                try
                {
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "1");
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, null);

                    var ex = Assert.Throws<InvalidOperationException>(() => AppSettings.GetDatabasePath());
                    Assert.Contains("FASTMAILSEARCH_REQUIRE_OVERRIDE=1", ex.Message);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, origDb);
                }
            }
        }

        [Fact]
        public void AppSettings_RequireOverride_ThrowsException_WhenExplicitPathMatchesProductionDatabase()
        {
            lock (_envLock)
            {
                string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
                string? origDb = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
                try
                {
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "true");
                    // Point explicitly to default path with different casing and trailing slash
                    string defaultPathWithSlash = AppSettings.GetDefaultDatabasePath().ToLowerInvariant() + @"\";
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, defaultPathWithSlash);

                    var ex = Assert.Throws<InvalidOperationException>(() => AppSettings.GetDatabasePath());
                    Assert.Contains("matches canonical production database", ex.Message);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, origDb);
                }
            }
        }

        [Fact]
        public void AppSettings_RequireOverride_Passes_WhenCustomPathResolved()
        {
            lock (_envLock)
            {
                string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
                string? origDb = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
                try
                {
                    string customPath = Path.Combine(Path.GetTempPath(), "FmsStagingDatabase.db");
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "1");
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, customPath);

                    string resolved = AppSettings.GetDatabasePath();

                    Assert.Equal(customPath, resolved);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, origDb);
                }
            }
        }

        [Fact]
        public void CommandLineParser_ParseCommandLineArgs_ValidatesDbPathAndLogDir()
        {
            string[] args = new[] { "--db-path", @"C:\Test Staging\test.db", "--log-dir", @"C:\Test Logs" };

            var parsed = CommandLineParser.ParseCommandLineArgs(args);

            Assert.False(parsed.HasInvalidArguments);
            Assert.Equal(@"C:\Test Staging\test.db", parsed.ExplicitDbPath);
            Assert.Equal(@"C:\Test Logs", parsed.ExplicitLogDir);
        }

        [Theory]
        [InlineData("--db-path")]
        [InlineData("--log-dir")]
        [InlineData("--unknown-flag")]
        public void CommandLineParser_ParseCommandLineArgs_RejectsMissingValueOrUnrecognizedArg(string invalidArg)
        {
            string[] args = new[] { invalidArg };

            var parsed = CommandLineParser.ParseCommandLineArgs(args);

            Assert.True(parsed.HasInvalidArguments);
            Assert.NotNull(parsed.ErrorMessage);
        }

        [Fact]
        public void CommandLineParser_ValidateStartupDatabasePath_PassesWhenNoExplicitArgSupplied()
        {
            lock (_envLock)
            {
                string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
                string? origDb = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
                try
                {
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, null);
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, null);

                    bool valid = CommandLineParser.ValidateStartupDatabasePath(null, out string resolvedDbPath, out string? error);

                    Assert.True(valid);
                    Assert.Null(error);
                    Assert.Equal(AppSettings.GetDefaultDatabasePath(), resolvedDbPath);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, origDb);
                }
            }
        }

        [Fact]
        public void CommandLineParser_ValidateStartupDatabasePath_PassesWhenExplicitMatchesResolved()
        {
            lock (_envLock)
            {
                string? origDb = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
                try
                {
                    string customPath = Path.Combine(Path.GetTempPath(), "FmsMatchTest.db");
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, customPath);

                    bool valid = CommandLineParser.ValidateStartupDatabasePath(customPath, out string resolvedDbPath, out string? error);

                    Assert.True(valid);
                    Assert.Null(error);
                    Assert.Equal(customPath, resolvedDbPath);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, origDb);
                }
            }
        }

        [Fact]
        public void CommandLineParser_ValidateStartupDatabasePath_RejectsMismatchBetweenExplicitAndResolved()
        {
            lock (_envLock)
            {
                string? origDb = Environment.GetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable);
                try
                {
                    string envPath = Path.Combine(Path.GetTempPath(), "FmsEnvDatabase.db");
                    string explicitPath = Path.Combine(Path.GetTempPath(), "FmsExplicitDatabase.db");
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, envPath);

                    bool valid = CommandLineParser.ValidateStartupDatabasePath(explicitPath, out string resolvedDbPath, out string? error);

                    Assert.False(valid);
                    Assert.NotNull(error);
                    Assert.Contains("Database path mismatch", error);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(AppSettings.DatabasePathEnvironmentVariable, origDb);
                }
            }
        }

#if NETFRAMEWORK
        [Fact]
        public void ThisAddIn_CreateUIProcessStartInfo_ConfiguresNoShellExecute_AndPropagatesArgsAndEnv()
        {
            lock (_envLock)
            {
                string? origLog = Environment.GetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable);
                string? origReq = Environment.GetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable);
                try
                {
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, @"C:\Staging Logs");
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, "1");

                    string uiExe = @"C:\TestApp\FastMailSearch.UI.exe";
                    string installDir = @"C:\TestApp";
                    string dbPath = @"C:\Test Staging\FastMailSearch_Staging.db";

                    ProcessStartInfo psi = ThisAddIn.CreateUIProcessStartInfo(uiExe, installDir, dbPath, @"C:\Staging Logs");

                    Assert.False(psi.UseShellExecute);
                    Assert.Equal(uiExe, psi.FileName);
                    Assert.Equal(installDir, psi.WorkingDirectory);
                    Assert.Contains(@"--db-path ""C:\Test Staging\FastMailSearch_Staging.db""", psi.Arguments);
                    Assert.Contains(@"--log-dir ""C:\Staging Logs""", psi.Arguments);
                    Assert.Equal(dbPath, psi.EnvironmentVariables[AppSettings.DatabasePathEnvironmentVariable]);
                    Assert.Equal(@"C:\Staging Logs", psi.EnvironmentVariables[Logger.LogDirectoryEnvironmentVariable]);
                    Assert.Equal("1", psi.EnvironmentVariables[AppSettings.RequireOverrideEnvironmentVariable]);
                }
                finally
                {
                    Environment.SetEnvironmentVariable(Logger.LogDirectoryEnvironmentVariable, origLog);
                    Environment.SetEnvironmentVariable(AppSettings.RequireOverrideEnvironmentVariable, origReq);
                }
            }
        }
#endif
    }
}
