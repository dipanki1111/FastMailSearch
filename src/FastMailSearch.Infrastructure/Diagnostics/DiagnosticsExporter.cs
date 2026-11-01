using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;

namespace FastMailSearch.Infrastructure.Diagnostics
{
    public class DiagnosticsExporter
    {
        private readonly EnvironmentDetector _detector;
        private readonly IIndexHealthService _healthService;
        private readonly IEmailRepository _repository;

        public DiagnosticsExporter(EnvironmentDetector detector, IIndexHealthService healthService, IEmailRepository repository)
        {
            _detector = detector;
            _healthService = healthService;
            _repository = repository;
        }

        public async System.Threading.Tasks.Task<string> ExportZipAsync(string targetFolder)
        {
            if (string.IsNullOrWhiteSpace(targetFolder) || !Directory.Exists(targetFolder))
            {
                targetFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FastMailSearch", "Diagnostics");
            }

            if (!Directory.Exists(targetFolder))
            {
                Directory.CreateDirectory(targetFolder);
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string zipFileName = $"FastMailSearch-Diagnostics-{timestamp}.zip";
            string zipFilePath = Path.Combine(targetFolder, zipFileName);

            string tempDir = Path.Combine(Path.GetTempPath(), $"FMS_Diag_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                // 1. Environment Diagnostics
                var env = _detector.DetectEnvironment();
                string envJson = SerializeEnvironmentDiagnostics(env);
                File.WriteAllText(Path.Combine(tempDir, "EnvironmentReport.json"), envJson, Encoding.UTF8);

                // 2. Health Summary
                var health = await _healthService.GetHealthSummaryAsync();
                string healthJson = SerializeHealthSummary(health);
                File.WriteAllText(Path.Combine(tempDir, "IndexHealthSummary.json"), healthJson, Encoding.UTF8);

                // 3. Database Statistics & Integrity Report
                var dbSize = await _repository.GetDatabaseSizeMbAsync();
                var integrity = await _repository.ScanIndexIntegrityAsync();
                var dbStats = new StringBuilder();
                dbStats.AppendLine("=== FASTMAIL SEARCH DATABASE STATISTICS ===");
                dbStats.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                dbStats.AppendLine($"Database Size: {dbSize:F2} MB");
                dbStats.AppendLine($"Integrity Result: {(integrity.IsHealthy ? "HEALTHY" : "CORRUPT/DEGRADED")}");
                dbStats.AppendLine($"Summary: {integrity.SummaryText}");
                dbStats.AppendLine($"Missing FTS Count: {integrity.MissingFtsCount}");
                dbStats.AppendLine($"Orphan FTS Count: {integrity.OrphanFtsCount}");
                dbStats.AppendLine($"Duplicate Count: {integrity.DuplicateCount}");
                File.WriteAllText(Path.Combine(tempDir, "DatabaseStatistics.txt"), dbStats.ToString(), Encoding.UTF8);

                // 4. Scrubbed Log Snippet
                string logDir = Logging.Logger.LogDirectory;
                string masterLog = Path.Combine(logDir, "FastMailSearch.log");
                if (File.Exists(masterLog))
                {
                    List<string> lines = new List<string>();
                    for (int retry = 0; retry < 3; retry++)
                    {
                        try
                        {
                            using (var fs = new FileStream(masterLog, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            using (var reader = new StreamReader(fs, Encoding.UTF8))
                            {
                                string? line;
                                while ((line = await reader.ReadLineAsync()) != null)
                                {
                                    lines.Add(line);
                                }
                            }
                            break;
                        }
                        catch (IOException)
                        {
                            if (retry == 2) throw;
                            await Task.Delay(50);
                        }
                    }

                    int take = Math.Min(500, lines.Count);
                    var snippet = new StringBuilder();
                    snippet.AppendLine("=== RECENT TECHNICAL LOGS (PII-SCRUBBED) ===");
                    for (int i = lines.Count - take; i < lines.Count; i++)
                    {
                        snippet.AppendLine(Logging.Logger.ScrubPii(lines[i]));
                    }
                    File.WriteAllText(Path.Combine(tempDir, "RecentTechnicalLogs.txt"), snippet.ToString(), Encoding.UTF8);
                }

                // Zip folder
                if (File.Exists(zipFilePath)) File.Delete(zipFilePath);
                ZipFile.CreateFromDirectory(tempDir, zipFilePath);

                return zipFilePath;
            }
            finally
            {
                if (Directory.Exists(tempDir))
                {
                    try { Directory.Delete(tempDir, true); } catch { }
                }
            }
        }

        private string SerializeEnvironmentDiagnostics(EnvironmentDiagnostics env)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine($"  \"OsVersion\": \"{Sanitize(env.OsVersion)}\",");
            sb.AppendLine($"  \"OsArchitecture\": \"{Sanitize(env.OsArchitecture)}\",");
            sb.AppendLine($"  \"OutlookVersion\": \"{Sanitize(env.OutlookVersion)}\",");
            sb.AppendLine($"  \"OutlookBitness\": \"{Sanitize(env.OutlookBitness)}\",");
            sb.AppendLine($"  \"OfficeVersion\": \"{Sanitize(env.OfficeVersion)}\",");
            sb.AppendLine($"  \"HasExchangeAccount\": {env.HasExchangeAccount.ToString().ToLower()},");
            sb.AppendLine($"  \"HasM365Account\": {env.HasM365Account.ToString().ToLower()},");
            sb.AppendLine($"  \"HasImapAccount\": {env.HasImapAccount.ToString().ToLower()},");
            sb.AppendLine($"  \"HasPopAccount\": {env.HasPopAccount.ToString().ToLower()},");
            sb.AppendLine($"  \"StoreCount\": {env.StoreCount},");
            sb.AppendLine($"  \"DefaultStoreType\": \"{Sanitize(env.DefaultStoreType)}\",");
            sb.AppendLine($"  \"FastMailSearchVersion\": \"{Sanitize(env.FastMailSearchVersion)}\",");
            sb.AppendLine($"  \"DatabasePath\": \"{Sanitize(env.DatabasePath)}\",");
            sb.AppendLine($"  \"DatabaseSizeMb\": {env.DatabaseSizeMb:F2},");
            sb.AppendLine($"  \"SqliteVersion\": \"{Sanitize(env.SqliteVersion)}\",");
            sb.AppendLine($"  \"IsFts5Available\": {env.IsFts5Available.ToString().ToLower()},");
            sb.AppendLine($"  \"Timestamp\": \"{env.Timestamp:o}\"");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private string SerializeHealthSummary(IndexHealthSummary h)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine($"  \"HealthState\": \"{h.HealthState}\",");
            sb.AppendLine($"  \"StatusState\": \"{Sanitize(h.StatusState)}\",");
            sb.AppendLine($"  \"IndexedEmailCount\": {h.IndexedEmailCount},");
            sb.AppendLine($"  \"PendingQueueCount\": {h.PendingQueueCount},");
            sb.AppendLine($"  \"FailedItemCount\": {h.FailedItemCount},");
            sb.AppendLine($"  \"RetryCount\": {h.RetryCount},");
            sb.AppendLine($"  \"IsSafeMode\": {h.IsSafeMode.ToString().ToLower()},");
            sb.AppendLine($"  \"SafeModeReason\": \"{Sanitize(h.SafeModeReason)}\",");
            sb.AppendLine($"  \"CurrentOperation\": \"{Sanitize(h.CurrentOperation ?? "Idle")}\",");
            sb.AppendLine($"  \"DatabaseSizeMb\": {h.DatabaseSizeMb:F2}");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private string Sanitize(string s) => s.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }
}
