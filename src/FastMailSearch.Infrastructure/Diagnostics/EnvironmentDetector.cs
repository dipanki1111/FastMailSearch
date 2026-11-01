using System;
using System.IO;
using System.Reflection;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Database;

namespace FastMailSearch.Infrastructure.Diagnostics
{
    public class EnvironmentDetector
    {
        private readonly IOutlookAdapter? _outlookAdapter;
        private readonly DatabaseConnectionFactory _dbFactory;

        public EnvironmentDetector(DatabaseConnectionFactory dbFactory, IOutlookAdapter? outlookAdapter = null)
        {
            _dbFactory = dbFactory;
            _outlookAdapter = outlookAdapter;
        }

        public EnvironmentDiagnostics DetectEnvironment(bool includeOutlookStores = true)
        {
            var diag = new EnvironmentDiagnostics
            {
                OsVersion = Environment.OSVersion.ToString(),
                OsArchitecture = Environment.Is64BitOperatingSystem ? "x64" : "x86",
                FastMailSearchVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0.0",
                AddInInstallationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty,
                DatabasePath = _dbFactory?.DbPath ?? string.Empty,
                Timestamp = DateTime.Now
            };

            // Database metrics & SQLite / FTS5 checks (fully independent of Outlook)
            try
            {
                if (_dbFactory != null && File.Exists(_dbFactory.DbPath))
                {
                    diag.DatabaseSizeMb = new FileInfo(_dbFactory.DbPath).Length / (1024.0 * 1024.0);
                }

                if (_dbFactory != null)
                {
                    using (var conn = _dbFactory.CreateConnection())
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "SELECT sqlite_version();";
                            var res = cmd.ExecuteScalar();
                            diag.SqliteVersion = res?.ToString() ?? "Unknown";
                        }

                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "SELECT count(*) FROM sqlite_master WHERE type='table' AND (name='EmailSearch' OR name='Emails_fts');";
                            var res = cmd.ExecuteScalar();
                            diag.IsFts5Available = res != null && Convert.ToInt32(res) > 0;
                        }
                    }
                }
            }
            catch
            {
                diag.SqliteVersion = "Unknown";
                diag.IsFts5Available = false;
            }

            diag.OutlookBitness = _outlookAdapter != null ? _outlookAdapter.GetOfficeBitness() : (Environment.Is64BitProcess ? "64-bit" : "32-bit");
            diag.OfficeVersion = $"Microsoft Office ({diag.OutlookBitness})";

            // If includeOutlookStores is false, return immediately without slow STA store enumeration
            if (!includeOutlookStores)
            {
                diag.OutlookVersion = "Connected";
                return diag;
            }

            // Outlook store metrics (only when explicitly requested)
            if (_outlookAdapter != null && _outlookAdapter.IsOutlookAvailable())
            {
                try
                {
                    diag.OutlookVersion = _outlookAdapter.GetOutlookVersion();

                    var mailboxes = _outlookAdapter.GetMailboxes();
                    diag.StoreCount = mailboxes?.Count ?? 0;

                    if (mailboxes != null && mailboxes.Count > 0)
                    {
                        foreach (var mb in mailboxes)
                        {
                            string name = mb.DisplayName ?? string.Empty;
                            if (mb.IsPst)
                            {
                                diag.HasPopAccount = true;
                            }
                            else if (name.Contains("Exchange") || name.Contains("@"))
                            {
                                diag.HasExchangeAccount = true;
                                diag.HasM365Account = true;
                            }
                            else
                            {
                                diag.HasImapAccount = true;
                            }
                        }

                        diag.DefaultStoreType = mailboxes[0].IsPst ? "PST / Local" : "Exchange / M365";
                    }
                }
                catch
                {
                    diag.OutlookVersion = "Unavailable";
                }
            }
            else
            {
                diag.OutlookVersion = "Not Connected";
            }

            return diag;
        }
    }
}
