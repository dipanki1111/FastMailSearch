using System;
using System.Collections.Generic;

namespace FastMailSearch.Core.Models
{
    public class EmailDocument
    {
        public long Id { get; set; }
        public string StoreId { get; set; } = string.Empty;
        public string EntryId { get; set; } = string.Empty;
        public string ConversationId { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string SenderName { get; set; } = string.Empty;
        public string SenderEmail { get; set; } = string.Empty;
        public string Recipients { get; set; } = string.Empty;
        public string Cc { get; set; } = string.Empty;
        public string Bcc { get; set; } = string.Empty;
        public string BodyText { get; set; } = string.Empty;
        public DateTime ReceivedDate { get; set; }
        public DateTime SentDate { get; set; }
        public string FolderName { get; set; } = string.Empty;
        public string FolderEntryId { get; set; } = string.Empty;
        public bool HasAttachment { get; set; }
        public bool IsRead { get; set; }
        public bool IsFlagged { get; set; }
        public int Importance { get; set; } // 0 = Low, 1 = Normal, 2 = High
        public string Categories { get; set; } = string.Empty;
        public long Size { get; set; }
        public DateTime LastModified { get; set; }
        public DateTime IndexedDate { get; set; }

        public List<EmailAttachment> Attachments { get; set; } = new List<EmailAttachment>();

        public string GetAttachmentNamesJoined()
        {
            if (Attachments == null || Attachments.Count == 0)
                return string.Empty;

            var names = new List<string>();
            foreach (var att in Attachments)
            {
                if (!string.IsNullOrWhiteSpace(att.Name))
                    names.Add(att.Name);
            }
            return string.Join(" ", names);
        }
    }
}
