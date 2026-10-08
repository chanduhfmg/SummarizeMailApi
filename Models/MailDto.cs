using SummarizeMail.Service;

namespace SummarizeMail.Models
{
    public class MailDto
    {
        // ───── Basic Info ─────
        public string MessageId { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;

        // ───── Sender / Recipients ─────
        public string From { get; set; } = string.Empty;
        public List<string> To { get; set; } = new();
        public List<string> Cc { get; set; } = new();

        // ───── Dates ─────
        public DateTime Received { get; set; }
        public DateTime Sent { get; set; }   // ✅ NEW
        public bool HasAttachments { get; set; }

        public bool IsRead { get;set; }

        // ───── Body ─────
        public string Body { get; set; } = string.Empty;         // HTML


        // ───── Attachments ─────
        public List<HelpdeskAttachment> Attachments { get; set; } = new();

        public string ConversationId { get; set; }
    }
}