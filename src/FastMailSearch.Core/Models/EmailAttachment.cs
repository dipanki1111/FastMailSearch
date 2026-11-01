namespace FastMailSearch.Core.Models
{
    public class EmailAttachment
    {
        public long Id { get; set; }
        public long EmailId { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Extension { get; set; } = string.Empty;
        public long Size { get; set; }
    }
}
