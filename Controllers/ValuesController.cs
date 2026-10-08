using Microsoft.AspNetCore.Mvc;
using Microsoft.Exchange.WebServices.Data;
using SummarizeMail.Models;
using SummarizeMail.Service;
using System;


// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace SummarizeMail.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        private readonly ExchangeMailService _mailService;

        public ValuesController(ExchangeMailService mailService)
        {
            _mailService = mailService;
        }

        //[HttpGet("inbox/{mail}/count{num}")]
        //public IActionResult GetInbox(string mail, int num)
        //{
        //    var mails = _mailService.GetInboxx(mail , num);

        //    return Ok(mails);
        //}

        [HttpGet("inbox/{mail}/count{num}")]
        public async Task<IActionResult> GetInbox(string mail, int num)
        {
            var mails = await _mailService.GetInboxx(mail, num);

            return Ok(mails);
        }

        [HttpGet("sentitems/{mail}/count{num}")]
        public async Task<IActionResult> GetSentItems(string mail, int num)
        {
            var mails = await _mailService.GetSentMailsForTracking(mail, num);
            return Ok(mails);
        }

        [HttpGet("sent7days")]
        public async Task<IActionResult> GetSevenDays(string email)
        {
            var mails = await _mailService.GetSentMailsOfSevenDays(email);
            return Ok(mails);
        }

        [HttpGet("onedaymails/{mail}")]
        public async Task<IActionResult> GetOneDayMails(string mail)
        {
            var mails = await _mailService.GetOneDayMails(mail);
            return Ok(mails);
        }

        [HttpGet("inbox/onemail")]
        public async Task<IActionResult> GetOneMail(string messageId, string targetUserEmail)
        {
            var mail = await _mailService.GetOneMail(messageId, targetUserEmail);
            return Ok(mail);
        }

        //to,cc and attachments list
        [HttpGet("mail/unread/count")]
        public async Task<IActionResult> GetUnreadMailCount(string targetUserEmail, int count)
        {
            var stats = await _mailService.GetMailStats(targetUserEmail, count);
            return Ok(stats);
        }

        //download attachment
        [HttpGet("attachments/download")]
        public async Task<IActionResult> DownloadAttachment(string messageId, string attachmentId, string targetUserEmail)
        {
            var attachmentContent = await _mailService.DownloadAttachment(messageId, attachmentId, targetUserEmail);
            return File(attachmentContent.Content, "application/octet-stream", attachmentContent.FileName);
        }


        // GET api/<ValuesController>/5
        [HttpPost("mail/read-status")]
        public async Task<IActionResult> UpdateReadStatus(
            string email,
            string messageId,
            bool isRead
        )
        {
            var result = await _mailService.SetMailReadStatus(email, messageId, isRead);

            if (!result)
                return BadRequest(result);

            return Ok(result);
        }


        [HttpPost("mail/send")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> SendEmail([FromForm] SendMailRequest request)
        {
            if (request.To == null || request.FromEmail == null)
                return BadRequest("At least one recipient is required");

            bool success = await _mailService.SendMailAsServiceAccount(
                request.To,
                request.Cc,
                request.Subject,
                request.BodyHtml,
                request.Attachments,
                request.FromEmail
            );

            return success ? Ok("Mail sent") : BadRequest("Mail failed");
        }

        [HttpPost("mail/SentAsRemainder")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> SendTrackingService([FromForm] SendMailAsRemainder request)
        {
            try
            {
                if (request.To == null || !request.To.Any())
                    return BadRequest("At least one recipient is required");

                bool success = await _mailService.SendMailTracking(
                    request.To,
                    request.Cc,
                    request.Subject,
                    request.BodyHtml,
                    request.Attachments
                );


                if (success)
                    return Ok("Mail sent");

                return BadRequest("Mail failed");
            }
            catch (Exception ex)
            {
                // Optional: log exception here
                //_logger.LogError(ex, "Error while sending tracking mail");

                return StatusCode(500, new
                {
                    Message = "An error occurred while sending mail",
                    Error = ex.Message,
                    InnerException = ex.InnerException?.ToString(),
                    Exception = ex.ToString()

                });
            }
        }


       //updated reply and replyall

        [HttpPost("mail/reply")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> ReplyToMail([FromForm] ReplyMailRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.MessageItemId))
                return BadRequest("MessageItemId is required");

            // ✅ Fix for Reply vs Reply All
            if (!request.IsReplyAll && (request.To == null || !request.To.Any()))
                return BadRequest("At least one TO recipient is required for reply");

            if (string.IsNullOrWhiteSpace(request.ReplyBodyHtml))
                return BadRequest("Reply body is required");

            bool success = await _mailService.ReplyToMail(
                request.MessageItemId,
                request.ReplyBodyHtml,
                request.To,
                request.Cc,
                request.Attachments,
                request.TargetUserEmail, // ✅ fixed
                request.IsReplyAll
            );

            if (!success)
                return StatusCode(500, "Failed to reply to email");

            return Ok(new { message = "Reply sent successfully" });
        }

        [HttpPost("forwardmail")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> ForwardEmail([FromForm] ForwardMailDto details, [FromForm] string targetUserEmail)
        {
            try
            {
                var result = await _mailService.ForwordEmail(details, targetUserEmail);

                if (!result.success)
                {
                    return StatusCode(500, new
                    {
                        success = false,
                        message = "Mail forwarding failed",
                        error = result.error
                    });
                }
                return Ok(new
                {
                    success = true,
                    message = "mail forwarded successfully",
                    details = details
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    error = ex.Message
                });
            }

        }


        [HttpGet("meetings")]
        public async Task<IActionResult> GetMeetings(
               [FromQuery] string email,
               [FromQuery] DateTime startDate,
               [FromQuery] DateTime endDate)
        {
            try
            {
                if (string.IsNullOrEmpty(email))
                    return BadRequest("Email is required.");

                var meetings = await _mailService
                    .GetCalendarMeetings(email, startDate, endDate);

                return Ok(meetings);
            }
            catch (ApplicationException ex)
            {
                return StatusCode(500, ex.Message);
            }
            catch (Exception)
            {
                return StatusCode(500, "Unexpected error occurred.");
            }

        }


        [HttpPost("meetings/Action")]
        public async Task<IActionResult> RespondMeeting(
    [FromQuery] string meetingId,
    [FromQuery] MeetingAction action,
    [FromQuery] string targetUserEmail)
        {
            try
            {
                Console.WriteLine($"from the controller {meetingId}");
                if (string.IsNullOrWhiteSpace(meetingId))
                    return BadRequest("MeetingId is required.");

                await _mailService.RespondToMeeting(meetingId, action, targetUserEmail);

                return Ok(new
                {
                    success = true,
                    message = $"Meeting {action} successful.",
                    meetingId = meetingId
                });
            }
            catch (ApplicationException ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    error = ex.Message
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    error = ex.Message
                });
            }
        }





        //get one mail

        [HttpGet("mail")]
        public IActionResult GetMail([FromQuery] string messageId, [FromQuery] string targetUserEmail)
        {
            try
            {
                var mail = _mailService.GetOneMail(messageId, targetUserEmail);

                return Ok(new
                {
                    success = true,
                    data = mail
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    error = ex.Message
                });
            }


        }

        [HttpGet("contacts/personal/{email}")]
        public async Task<IActionResult> SearchPersonalContacts(string email, [FromQuery] string q)
        {
            if (string.IsNullOrWhiteSpace(q) || q.Length < 2)
                return Ok(new List<object>());

            var result = await _mailService.SearchPersonalContacts(email, q);

            if (!result.Success)
            {
                // 502 signals "upstream Graph call failed" distinctly from
                // "call succeeded, genuinely no matches" (which stays 200/[])
                return StatusCode(502, new
                {
                    error = result.ErrorCode,
                    message = result.ErrorMessage
                });
            }

            return Ok(result.Contacts);
        }


        //[HttpDelete("deleted-items/{messageId}")]
        //public async Task<IActionResult> DeleteMailFromDeletedItems(string messageId, [FromQuery] string targetUserEmail)
        //{
        //    var success = await _mailService.DeleteMailFromDeletedItems(messageId, targetUserEmail);

        //    if (!success)
        //        return StatusCode(500, "Failed to permanently delete mail.");

        //    return NoContent();
        //}

        [HttpGet("health")]
        public IActionResult Health()
        {
            return Ok(new
            {
                status = "ok",
                service = "mail-service"
            });
        }


        public class SendMailRequest
        {
            public List<string> To { get; set; } = new();
            public List<string>? Cc { get; set; }
            public string Subject { get; set; } = string.Empty;
            public string BodyHtml { get; set; } = string.Empty;

            public List<IFormFile>? Attachments { get; set; }
            public string? FromEmail { get; set; }


        }

        public class SendMailAsRemainder
        {
            public List<string> To { get; set; } = new();
            public List<string>? Cc { get; set; }
            public string Subject { get; set; } = string.Empty;
            public string BodyHtml { get; set; } = string.Empty;

            public List<IFormFile>? Attachments { get; set; }


        }



        public class ReplyMailRequest
        {
            public string MessageItemId { get; set; } = null!;
            public string ReplyBodyHtml { get; set; } = null!;
            public List<string> To { get; set; } = new();
            public List<string>? Cc { get; set; }

            public List<IFormFile>? Attachments { get; set; }
            public string TargetUserEmail { get; set; }

            public bool IsReplyAll { get; set; }
        }


    }
}
public enum MeetingAction
{
    Accept,
    Reject
}
