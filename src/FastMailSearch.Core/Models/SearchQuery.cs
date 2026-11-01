using System;
using System.Collections.Generic;

namespace FastMailSearch.Core.Models
{
    public enum SortOption
    {
        Relevance,
        NewestFirst,
        OldestFirst
    }

    public class SearchQuery
    {
        public string RawQuery { get; set; } = string.Empty;
        public string FreeText { get; set; } = string.Empty;
        public List<string> ExactPhrases { get; set; } = new List<string>();
        public List<string> TextTerms { get; set; } = new List<string>();
        public List<string> IncludeTerms { get; set; } = new List<string>();
        public List<string> ExcludeTerms { get; set; } = new List<string>();
        public string SenderFilter { get; set; } = string.Empty;
        public string RecipientFilter { get; set; } = string.Empty;
        public string CcFilter { get; set; } = string.Empty;
        public string SubjectFilter { get; set; } = string.Empty;
        public string FilenameFilter { get; set; } = string.Empty;
        public string FolderFilter { get; set; } = string.Empty;
        public string StoreFilter { get; set; } = string.Empty;
        public bool? HasAttachmentFilter { get; set; }
        public bool? IsReadFilter { get; set; }
        public bool? IsFlaggedFilter { get; set; }
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public string BooleanOperator { get; set; } = "AND";
        public SortOption SortOrder { get; set; } = SortOption.Relevance;
        public bool IsInvalidSyntax { get; set; }
        public string SyntaxErrorMessage { get; set; } = string.Empty;
        public int Limit { get; set; } = 50;
        public int Offset { get; set; } = 0;

        public int PageSize
        {
            get => Limit > 0 ? Limit : 50;
            set
            {
                int sz = value < 1 ? 50 : value;
                int currentP = Page;
                Limit = sz;
                Offset = (currentP - 1) * Limit;
            }
        }

        public int Page
        {
            get => Limit > 0 ? (Offset / Limit) + 1 : 1;
            set
            {
                int p = value < 1 ? 1 : value;
                int sz = Limit > 0 ? Limit : 50;
                Offset = (p - 1) * sz;
            }
        }

        // Property aliases to support explicit Phase 1 query object structure
        public string From { get => SenderFilter; set => SenderFilter = value; }
        public string To { get => RecipientFilter; set => RecipientFilter = value; }
        public string Cc { get => CcFilter; set => CcFilter = value; }
        public string Subject { get => SubjectFilter; set => SubjectFilter = value; }
        public string Filename { get => FilenameFilter; set => FilenameFilter = value; }
        public string Folder { get => FolderFilter; set => FolderFilter = value; }
        public bool? HasAttachment { get => HasAttachmentFilter; set => HasAttachmentFilter = value; }
        public DateTime? AfterDate { get => DateFrom; set => DateFrom = value; }
        public DateTime? BeforeDate { get => DateTo; set => DateTo = value; }
    }
}
