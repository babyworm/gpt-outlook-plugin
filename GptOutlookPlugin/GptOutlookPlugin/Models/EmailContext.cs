namespace GptOutlookPlugin.Models
{
    public class EmailContext
    {
        public string Subject { get; set; } = "";
        public string Body { get; set; } = "";
        public string BodyHtml { get; set; } = "";
        public string Recipients { get; set; } = "";
        public string SenderEmail { get; set; } = "";
        public bool IsComposing { get; set; }
    }
}
