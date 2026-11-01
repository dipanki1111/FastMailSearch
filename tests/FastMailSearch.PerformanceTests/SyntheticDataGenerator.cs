using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;

namespace FastMailSearch.PerformanceTests
{
    public static class SyntheticDataGenerator
    {
        private static readonly string[] Senders = new[] { "ABC Ltd", "John Smith", "Accounts", "Support", "Jane Doe", "Finance Dept", "Global Tech", "Billing Office", "John Example", "KFintech Corp" };
        private static readonly string[] Domains = new[] { "abc.com", "example.com", "global.org", "vendor.net", "corp.com", "finances.io", "kfintech.com" };
        private static readonly string[] Subjects = new[] {
            "Invoice INV-2026-001", "RE: Invoice ABC", "Payment Received", "Weekly Status Report",
            "Urgent: Contract Revision", "Quarterly Financial Overview", "project meeting notes",
            "Purchase Order PO-9921", "Tax statement 2026", "Security Update Notification", "KFintech Invoice Statement"
        };
        private static readonly string[] Folders = new[] { "Inbox", "Sent Items", "Archive", "Projects", "Finance" };
        private static readonly string[] AttachmentFiles = new[] { "invoice.pdf", "statement.xlsx", "contract.docx", "report.pdf", "receipt.png", "" };

        public static async Task GenerateSyntheticEmailsAsync(IEmailRepository repository, int count, Action<int>? progressCallback = null)
        {
            var rnd = new Random(42); // deterministic seed for reproducibility
            int batchSize = 1000;
            var batch = new List<EmailDocument>();

            for (int i = 1; i <= count; i++)
            {
                string senderName = Senders[rnd.Next(Senders.Length)];
                string domain = Domains[rnd.Next(Domains.Length)];
                string senderEmail = $"{senderName.Replace(" ", ".").ToLowerInvariant()}@{domain}";
                string subject = Subjects[rnd.Next(Subjects.Length)] + $" #{i}";
                string folder = Folders[rnd.Next(Folders.Length)];
                string attName = AttachmentFiles[rnd.Next(AttachmentFiles.Length)];
                bool hasAtt = !string.IsNullOrEmpty(attName);

                DateTime receivedDate = DateTime.Now.AddDays(-rnd.Next(0, 365)).AddMinutes(-rnd.Next(0, 1440));

                var doc = new EmailDocument
                {
                    StoreId = "BenchmarkStore",
                    EntryId = $"BenchEntry_{i:D8}",
                    ConversationId = $"Conv_{i % 500}",
                    Subject = subject,
                    SenderName = senderName,
                    SenderEmail = senderEmail,
                    Recipients = "user@company.com",
                    BodyText = $"Hello, this is synthetic benchmark email body content number {i}. Contains reference terms like invoice ABC, payment 2026, and project status details.",
                    ReceivedDate = receivedDate,
                    SentDate = receivedDate.AddMinutes(-5),
                    FolderName = folder,
                    FolderEntryId = $"FolderEntry_{folder}",
                    HasAttachment = hasAtt,
                    IsRead = rnd.Next(0, 2) == 1,
                    IsFlagged = rnd.Next(0, 5) == 0,
                    Importance = rnd.Next(0, 3),
                    Size = rnd.Next(1024, 102400),
                    LastModified = receivedDate,
                    IndexedDate = DateTime.Now
                };

                if (hasAtt)
                {
                    doc.Attachments.Add(new EmailAttachment
                    {
                        Name = attName,
                        Extension = System.IO.Path.GetExtension(attName),
                        Size = rnd.Next(5000, 500000)
                    });
                }

                batch.Add(doc);

                if (batch.Count >= batchSize || i == count)
                {
                    await repository.SaveEmailBatchAsync(batch);
                    batch.Clear();
                    progressCallback?.Invoke(i);
                }
            }
        }
    }
}
