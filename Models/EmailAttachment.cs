namespace OutlookAddIn.Models
{
    public class EmailAttachment
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string ContentId { get; set; }
        public bool IsInline { get; set; }
    }
}
