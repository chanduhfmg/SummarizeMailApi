namespace SummarizeMail.Models
{
    public class CalendarDto
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Location { get; set; }
        public string Organizer { get; set; }
        public List<string> RequiredAttendees { get; set; } = new();
        public List<string> OptionalAttendees { get; set; } = new();
        public bool IsAllDay { get; set; }
        public bool IsCancelled { get; set; }
        public string MeetingId { get; set; }

       public string ResponseStatus { get; set; }

    }
}
