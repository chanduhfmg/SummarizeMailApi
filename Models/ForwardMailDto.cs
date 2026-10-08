namespace SummarizeMail.Models
{
    public class ForwardMailDto
    {
        public string? Subject {  get; set; }
        public string? Body { get; set; }    

        public string MessageItemId {  get; set; }

        public List<string>? ToList { get; set; } = new();

        public List<string>? CcList { get; set; } = new();

        public List<IFormFile>? Attachments { get; set; }


    }
}
