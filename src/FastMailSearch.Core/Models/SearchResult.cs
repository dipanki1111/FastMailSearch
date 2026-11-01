using System;
using System.Collections.Generic;

namespace FastMailSearch.Core.Models
{
    public class SearchResultItem
    {
        public long EmailId { get; set; }
        public string StoreId { get; set; } = string.Empty;
        public string EntryId { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string SenderName { get; set; } = string.Empty;
        public string SenderEmail { get; set; } = string.Empty;
        public string Recipients { get; set; } = string.Empty;
        public DateTime ReceivedDate { get; set; }
        public string FolderName { get; set; } = string.Empty;
        public string BodySnippet { get; set; } = string.Empty;
        public bool HasAttachment { get; set; }
        public bool IsRead { get; set; }
        public bool IsFlagged { get; set; }
        public int Importance { get; set; }
        public double Rank { get; set; }
        public List<string> AttachmentNames { get; set; } = new List<string>();

        public string DisplaySender => string.IsNullOrWhiteSpace(SenderName) ? SenderEmail : $"{SenderName} <{SenderEmail}>";
        public string DisplayDate => ReceivedDate.ToString("dd MMM yyyy hh:mm tt");
        public string ImportanceIcon => Importance == 2 ? "❗ High" : Importance == 0 ? "⬇ Low" : "";
        public string AttachmentIcon => HasAttachment ? "📎" : "";
        public string FlagIcon => IsFlagged ? "🚩" : "";
    }

    public class SearchResultContainer
    {
        public List<SearchResultItem> Items { get; set; } = new List<SearchResultItem>();
        public int TotalCount { get; set; }
        public int Page { get; set; } = 1;
        public int PageSize { get; set; } = 50;
        public int TotalPages => PageSize > 0 && TotalCount > 0 ? (int)Math.Ceiling((double)TotalCount / PageSize) : 0;
        public double ElapsedSeconds { get; set; }
        public long SearchTimeMs
        {
            get => (long)Math.Round(ElapsedSeconds * 1000.0);
            set => ElapsedSeconds = value / 1000.0;
        }
        public bool HasMoreResults { get; set; }
        public bool HasPreviousPage => Page > 1;
        public bool HasNextPage => Page < TotalPages;

        public PagedSearchResult<SearchResultItem> ToPagedResult()
        {
            return new PagedSearchResult<SearchResultItem>
            {
                Items = Items,
                TotalCount = TotalCount,
                Page = Page,
                PageSize = PageSize,
                ElapsedSeconds = ElapsedSeconds
            };
        }
    }

    public class PagedSearchResult<T>
    {
        public List<T> Items { get; set; } = new List<T>();
        public int TotalCount { get; set; }
        public int Page { get; set; } = 1;
        public int PageSize { get; set; } = 50;
        public int TotalPages => PageSize > 0 && TotalCount > 0 ? (int)Math.Ceiling((double)TotalCount / PageSize) : 0;
        public double ElapsedSeconds { get; set; }
        public long SearchTimeMs
        {
            get => (long)Math.Round(ElapsedSeconds * 1000.0);
            set => ElapsedSeconds = value / 1000.0;
        }
        public bool HasPreviousPage => Page > 1;
        public bool HasNextPage => Page < TotalPages;
    }
}
