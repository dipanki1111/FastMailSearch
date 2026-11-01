namespace FastMailSearch.Core.Models
{
    public class MailboxFolder
    {
        public long Id { get; set; }
        public string StoreId { get; set; } = string.Empty;
        public string EntryId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string ParentId { get; set; } = string.Empty;
        public int ItemCount { get; set; }
        public bool IsSelectedForIndexing { get; set; } = true;
    }

    public class MailboxInfo
    {
        public string StoreId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public bool IsPst { get; set; }
        public bool IsSelected { get; set; } = true;
    }
}
