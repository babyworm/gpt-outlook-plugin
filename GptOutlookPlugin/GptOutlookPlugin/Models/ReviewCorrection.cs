namespace GptOutlookPlugin.Models
{
    public class ReviewCorrection
    {
        public string Original { get; set; }
        public string Corrected { get; set; }
        public string Reason { get; set; }
        public bool Accepted { get; set; }
        public bool Skipped { get; set; }
    }
}
