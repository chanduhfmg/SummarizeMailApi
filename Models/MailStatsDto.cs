namespace SummarizeMail.Models
{
    public class MailStatsDto
    {
        public int UnreadCount { get; set; }
        public int TotalAttachments { get; set; }
        public int TotalToRecipients { get; set; }
        public int TotalCcRecipients { get; set; }

        public Dictionary<string, int> FileTypeCounts { get; set; } = new();
    }
}
