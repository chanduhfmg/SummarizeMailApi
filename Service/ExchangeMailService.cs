using HtmlAgilityPack;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
// Aliases for EWS types that share a name with a Microsoft.Graph.Models type.
// Both namespaces define "Mailbox" and "FileAttachment" -- with both usings
// present, the bare names are ambiguous. These aliases resolve it without
// having to fully-qualify every EWS call site. Remove each alias (and its
// usages below) as the methods that need it get converted to Graph and no
// longer reference the EWS version.
using EwsMailbox = Microsoft.Exchange.WebServices.Data.Mailbox;
using EwsFileAttachment = Microsoft.Exchange.WebServices.Data.FileAttachment;
using EwsBodyType = Microsoft.Exchange.WebServices.Data.BodyType;
using Org.BouncyCastle.Utilities;
using SummarizeMail.Models;
using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using static MailKit.Net.Imap.ImapMailboxFilter;
using static Org.BouncyCastle.Math.EC.ECCurve;
using GraphAttachment = Microsoft.Graph.Models.Attachment;
using FileAttachment = Microsoft.Graph.Models.FileAttachment;
using Microsoft.Graph.Users.Item.SendMail;
using BodyType = Microsoft.Graph.Models.BodyType;
using EmailAddress = Microsoft.Graph.Models.EmailAddress;
using Attendee = Microsoft.Graph.Models.Attendee;
using GraphDateTimeTimeZone = Microsoft.Graph.Models.DateTimeTimeZone;
using Microsoft.Graph.Users.Item.Events.Item.Accept;
using Microsoft.Graph.Users.Item.Events.Item.Decline;
using TimeZoneConverter;
using Microsoft.Graph.Users.Item.Messages.Item.CreateForward;
using Microsoft.Graph.Users.Item.CalendarView;
using Calendar = Microsoft.Graph.Models.Calendar;
using ContactFolder = Microsoft.Graph.Models.ContactFolder;
using Contact = Microsoft.Graph.Models.Contact;

namespace SummarizeMail.Service
{
    public class HelpdeskAttachment
    {
        public string Name { get; set; }
        public string Type { get; set; }
        //public string Hash { get; set; }
        public int Size { get; set; }
        public string AttachmentId { get; set; }

    }

    public class PagedMailResponse
    {
        public List<MailDto> Mails { get; set; }
        public int CurrentPage { get; set; }
        public int PageSize { get; set; }
        public bool HasMore { get; set; }
    }

    public class DownloadedAttachment
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public byte[] Content { get; set; }




        public string ContentType { get; set; }

        public string attachmentId { get; set; }
    }

    public class ExchangeMailService
    {
        private readonly GraphServiceClient _graphClient;
        private  readonly string _reminderMailbox;
        
        public ExchangeMailService(IGraphClientFactory graphClientFactory)
        {
            _graphClient = graphClientFactory.GetClient();
            _reminderMailbox = graphClientFactory.ReminderMailbox;
           

        }

        string GetPlainTextFromHtml(string html)
        {
            if (string.IsNullOrEmpty(html)) return string.Empty;

            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            return doc.DocumentNode.InnerText
                      .Replace("\r", "")
                      .Replace("\n\n", "\n")
                      .Trim();
        }
        // ── GRAPH VERSION ───────────────────────────────────────────────────────
        // Two changes from the EWS version:
        //   1. Auth: no per-call ExchangeService/creds — reuses the singleton
        //      GraphServiceClient from IGraphClientFactory.
        //   2. Payload: this now returns METADATA ONLY (id, subject, from, date,
        //      isRead, hasAttachments, conversationId + Graph's free bodyPreview
        //      snippet). Full HTML body and real attachment list are NOT loaded
        //      here — call GetOneMail for those once a specific message is opened.
        //      That's what keeps this fast even with large pageSize.
        public async Task<PagedMailResponse> GetInboxx(string targetUserEmail, int pageSize, int page = 1)
        {
            try
            {
                int skip = (page - 1) * pageSize;

                var result = await _graphClient.Users[targetUserEmail]
                    .MailFolders["Inbox"]
                    .Messages
                    .GetAsync(requestConfig =>
                    {
                        // NOTE: receivedDateTime must appear in $filter (and before isRead)
                        // because it's also used in $orderby. Graph rejects filter+orderby
                        // combinations where the sort property isn't in the filter first —
                        // you'd get "The restriction or sort order is too complex for this
                        // operation." The ge-1900 clause is a harmless always-true condition
                        // that exists purely to satisfy that rule.
                        requestConfig.QueryParameters.Filter =
                            "receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false";
                        requestConfig.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                        requestConfig.QueryParameters.Select = new[]
                        {
                            "id", "subject", "receivedDateTime", "from",
                            "isRead", "hasAttachments", "conversationId", "bodyPreview"
                        };
                        requestConfig.QueryParameters.Top = pageSize;
                        requestConfig.QueryParameters.Skip = skip;

                        // Keeps message IDs stable across mailbox moves (per team lead's note)
                        requestConfig.Headers.Add("Prefer", "IdType=\"ImmutableId\"");
                    });

                var mails = (result?.Value ?? new List<Message>())
                    .Select(item => new MailDto
                    {
                        From = item.From?.EmailAddress?.Address ?? string.Empty,
                        Subject = item.Subject ?? string.Empty,
                        Received = item.ReceivedDateTime?.UtcDateTime ?? default,
                        IsRead = item.IsRead ?? false,
                        MessageId = item.Id ?? string.Empty,
                        // Preview text only at list time -- see note above.
                        Body = item.BodyPreview ?? string.Empty,
                        HasAttachments = item.HasAttachments ?? false,
                        ConversationId = item.ConversationId ?? string.Empty,
                        Attachments = new List<HelpdeskAttachment>()
                    })
                    .ToList();

                return new PagedMailResponse
                {
                    Mails = mails,
                    CurrentPage = page,
                    PageSize = pageSize,
                    HasMore = !string.IsNullOrEmpty(result?.OdataNextLink)
                };
            }
            catch (ODataError odErr) when (odErr.ResponseStatusCode == 403)
            {
                // Expected per team lead's email if this mailbox isn't yet on the
                // Exchange Online Application Access Policy allowlist.
                throw new ApplicationException(
                    $"Graph returned 403 for {targetUserEmail} — this mailbox is likely " +
                    "not yet included in the Application Access Policy allowlist.", odErr);
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error fetching inbox.", ex);
            }
        }



        // GRAPH VERSION: internetMessageId is now a first-class field (no MAPI
        // extended property needed) for matching "is this a reply about sent-message
        // X." In-Reply-To still has to come from internetMessageHeaders on both sent
        // and inbox items, same as EWS's own header-fallback path -- Graph has no
        // first-class field for it. Inbox pagination is $skip-based (see GetInboxx)
        // instead of EWS's manual Offset loop.
        // (internetMessageId was being loaded on inbox items in the original method
        // too, but never actually read there -- dropped rather than carried over.)
        public async Task<List<SentMailTrackDto>> GetSentMailsForTracking(string targetUserEmail, int num)
        {
            try
            {
                DateTime startDate = DateTime.UtcNow.AddDays(-5);
                var startDateIso = startDate.ToString("yyyy-MM-ddTHH:mm:ssZ");

                // ── 1. Fetch Sent Items ──────────────────────────────────────────
                var sentResult = await _graphClient.Users[targetUserEmail].MailFolders["SentItems"].Messages
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Filter = $"sentDateTime ge {startDateIso}";
                        requestConfig.QueryParameters.Orderby = new[] { "sentDateTime desc" };
                        requestConfig.QueryParameters.Select = new[]
                        {
                            "id", "conversationId", "subject", "internetMessageId",
                            "internetMessageHeaders", "body", "toRecipients", "ccRecipients", "sentDateTime"
                        };
                        requestConfig.QueryParameters.Top = num;
                        requestConfig.Headers.Add("Prefer", "IdType=\"ImmutableId\"");
                    });

                var sentItems = sentResult?.Value ?? new List<Message>();

                // ── 2. Fetch Inbox (paginated) and build the reply lookup ────────
                var replyMap = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

                const int pageSize = 500;
                int skip = 0;
                List<Message> inboxPage;

                do
                {
                    var page = await _graphClient.Users[targetUserEmail].MailFolders["Inbox"].Messages
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.Filter = $"receivedDateTime ge {startDateIso}";
                            requestConfig.QueryParameters.Select = new[] { "id", "from", "internetMessageHeaders" };
                            requestConfig.QueryParameters.Top = pageSize;
                            requestConfig.QueryParameters.Skip = skip;
                        });

                    inboxPage = page?.Value ?? new List<Message>();

                    foreach (var inboxMail in inboxPage)
                    {
                        string inReplyToValue = inboxMail.InternetMessageHeaders?
                            .FirstOrDefault(h => string.Equals(h.Name, "In-Reply-To", StringComparison.OrdinalIgnoreCase))
                            ?.Value ?? string.Empty;

                        if (string.IsNullOrEmpty(inReplyToValue)) continue;

                        string fromAddress = inboxMail.From?.EmailAddress?.Address ?? string.Empty;
                        if (string.IsNullOrEmpty(fromAddress)) continue;

                        inReplyToValue = NormalizeMessageId(inReplyToValue);
                        if (string.IsNullOrEmpty(inReplyToValue)) continue;

                        if (!replyMap.ContainsKey(inReplyToValue))
                            replyMap[inReplyToValue] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        replyMap[inReplyToValue].Add(fromAddress);
                    }

                    skip += pageSize;
                }
                while (inboxPage.Count == pageSize);

                // ── 3. Build result DTOs ─────────────────────────────────────────
                var mails = new List<SentMailTrackDto>();

                foreach (var sentMail in sentItems)
                {
                    string sentInReplyToValue = sentMail.InternetMessageHeaders?
                        .FirstOrDefault(h => string.Equals(h.Name, "In-Reply-To", StringComparison.OrdinalIgnoreCase))
                        ?.Value ?? string.Empty;

                    bool isReply = !string.IsNullOrWhiteSpace(sentInReplyToValue);

                    bool hasQuestionMark = false;
                    if (isReply)
                    {
                        string emailBody = sentMail.Body?.Content ?? string.Empty;
                        hasQuestionMark = ContainsQuestionMark(emailBody);
                    }

                    string originalMessageId = NormalizeMessageId(sentMail.InternetMessageId ?? string.Empty);

                    var repliedByAddresses = replyMap.TryGetValue(originalMessageId, out var set)
                        ? set
                        : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    var allRecipients = (sentMail.ToRecipients ?? new List<Recipient>())
                        .Select(r => r.EmailAddress?.Address ?? string.Empty)
                        .Where(r => !string.IsNullOrEmpty(r))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var notRepliedAddresses = allRecipients
                        .Where(r => !repliedByAddresses.Contains(r))
                        .ToList();

                    bool hasNonRepliers = notRepliedAddresses.Count > 0;

                    mails.Add(new SentMailTrackDto
                    {
                        MessageId = sentMail.Id ?? string.Empty,
                        ConversationId = sentMail.ConversationId ?? string.Empty,
                        To = (sentMail.ToRecipients ?? new List<Recipient>()).Select(r => r.EmailAddress?.Address ?? string.Empty).ToList(),
                        Cc = (sentMail.CcRecipients ?? new List<Recipient>()).Select(r => r.EmailAddress?.Address ?? string.Empty).ToList(),
                        SentAt = sentMail.SentDateTime?.UtcDateTime ?? default,
                        Subject = sentMail.Subject ?? string.Empty,
                        IsReply = isReply,
                        RepliedByAddresses = repliedByAddresses,
                        NotRepliedAddresses = notRepliedAddresses,
                        HasPendingReplies = hasNonRepliers,
                        HasQuestionMark = hasQuestionMark,
                        Body = sentMail.Body?.Content ?? string.Empty
                    });
                }

                return mails;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error fetching sent mails: " + ex.Message);
                throw;
            }
        }

        // ✅ NEW: Helper method to check for question mark
        private bool ContainsQuestionMark(string emailBody)
        {
            if (string.IsNullOrEmpty(emailBody))
                return false;

            // Remove HTML tags
            string plainText = Regex.Replace(emailBody, "<[^>]*>", " ");

            // Remove quoted text (lines starting with >)
            plainText = Regex.Replace(plainText, @"^>+.*$", string.Empty, RegexOptions.Multiline);

            // Check if contains "?"
            return plainText.Contains("?");
        }
        private static string NormalizeMessageId(string id) =>
            id?.Trim().Trim('<', '>').Trim() ?? string.Empty;


        // ── Updated DTO ──────────────────────────────────────────────────────────────
        public class SentMailTrackDto
        {
            public string MessageId { get; set; } = string.Empty;
            public string ConversationId { get; set; } = string.Empty;
            public List<string> To { get; set; } = new();
            public List<string> Cc { get; set; } = new();
            public DateTime SentAt { get; set; }
            public bool IsReply { get; set; }

            // ✅ FIX #2: Renamed from InReplyToMessageIds — this holds replier email addresses
            public HashSet<string> RepliedByAddresses { get; set; } = new();
            public string Subject { get; set; } = string.Empty;

            // ✅ Who has NOT replied yet
            public List<string> NotRepliedAddresses { get; set; } = new();

            // ✅ True if any recipient is still pending a reply
            public bool HasPendingReplies { get; set; }

            // ✅ NEW FIELD: Has question mark in body (only set if IsReply == true)
            public bool HasQuestionMark { get; set; }
            public string Body { get; set; }

            public string From { get; set; }

            public bool IsRead { get; set; }

        }

        // GRAPH VERSION: body is fetched as plain text via the
        // "outlook.body-content-type: text" Prefer header, matching this method's
        // original explicit RequestedBodyType = BodyType.Text (every other converted
        // method in this file kept HTML; this one specifically didn't, so that's
        // preserved here). internetMessageId is now first-class; In-Reply-To still
        // comes from internetMessageHeaders on both sent and inbox items -- same as
        // EWS's own fallback path. Inbox pagination is $skip-based instead of EWS's
        // manual Offset loop.
        public async Task<List<SentMailTrackDto>> GetSentMailsOfSevenDays(string targetUserEmail)
        {
            try
            {
                DateTime startDate = DateTime.UtcNow.AddDays(-7);
                var startDateIso = startDate.ToString("yyyy-MM-ddTHH:mm:ssZ");

                // ── Step 1: Fetch sent items ──────────────────────────────────────
                var sentResult = await _graphClient.Users[targetUserEmail].MailFolders["SentItems"].Messages
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Filter = $"sentDateTime ge {startDateIso}";
                        requestConfig.QueryParameters.Orderby = new[] { "sentDateTime desc" };
                        requestConfig.QueryParameters.Select = new[]
                        {
                            "id", "conversationId", "subject", "toRecipients", "body",
                            "internetMessageId", "internetMessageHeaders", "sentDateTime"
                        };
                        requestConfig.QueryParameters.Top = 1000;
                        requestConfig.Headers.Add("Prefer", "IdType=\"ImmutableId\"");
                        requestConfig.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                    });

                var sentItems = sentResult?.Value ?? new List<Message>();

                if (!sentItems.Any())
                    return new List<SentMailTrackDto>();

                // ── Step 2: Build DTOs, indexed both ways for the inbox pass below ─
                var mails = new List<SentMailTrackDto>();
                var sentMailIndex = new Dictionary<string, SentMailTrackDto>(StringComparer.OrdinalIgnoreCase);
                var conversationIndex = new Dictionary<string, SentMailTrackDto>(StringComparer.OrdinalIgnoreCase);

                foreach (var sentMail in sentItems)
                {
                    string sentInReplyToValue = sentMail.InternetMessageHeaders?
                        .FirstOrDefault(h => string.Equals(h.Name, "In-Reply-To", StringComparison.OrdinalIgnoreCase))
                        ?.Value ?? string.Empty;

                    bool isReply = !string.IsNullOrWhiteSpace(sentInReplyToValue);
                    string originalMessageId = NormalizeMessageId(sentMail.InternetMessageId ?? string.Empty);

                    var allRecipients = (sentMail.ToRecipients ?? new List<Recipient>())
                        .Select(r => r.EmailAddress?.Address ?? string.Empty)
                        .Where(r => !string.IsNullOrEmpty(r))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var dto = new SentMailTrackDto
                    {
                        MessageId = sentMail.Id ?? string.Empty,
                        ConversationId = sentMail.ConversationId ?? string.Empty,
                        To = allRecipients,
                        Cc = new List<string>(),
                        SentAt = sentMail.SentDateTime?.UtcDateTime ?? default,
                        Subject = sentMail.Subject ?? string.Empty,
                        IsReply = isReply,
                        RepliedByAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                        NotRepliedAddresses = allRecipients.ToList(),
                        HasPendingReplies = allRecipients.Count > 0,
                        HasQuestionMark = false,
                        Body = sentMail.Body?.Content ?? string.Empty
                    };

                    mails.Add(dto);

                    if (!string.IsNullOrEmpty(originalMessageId))
                        sentMailIndex[originalMessageId] = dto;

                    if (!string.IsNullOrEmpty(dto.ConversationId))
                        conversationIndex[dto.ConversationId] = dto;
                }

                // ── Step 3: Fetch Inbox (wider window, paginated) and match replies ─
                var inboxStartIso = startDate.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
                const int pageSize = 1000;
                int skip = 0;
                List<Message> inboxPage;

                do
                {
                    var page = await _graphClient.Users[targetUserEmail].MailFolders["Inbox"].Messages
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.Filter = $"receivedDateTime ge {inboxStartIso}";
                            requestConfig.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                            requestConfig.QueryParameters.Select = new[] { "id", "from", "conversationId", "internetMessageHeaders" };
                            requestConfig.QueryParameters.Top = pageSize;
                            requestConfig.QueryParameters.Skip = skip;
                        });

                    inboxPage = page?.Value ?? new List<Message>();

                    foreach (var inboxMail in inboxPage)
                    {
                        string inReplyToValue = NormalizeMessageId(inboxMail.InternetMessageHeaders?
                            .FirstOrDefault(h => string.Equals(h.Name, "In-Reply-To", StringComparison.OrdinalIgnoreCase))
                            ?.Value ?? string.Empty);

                        if (string.IsNullOrEmpty(inReplyToValue)) continue;

                        SentMailTrackDto matchedSent = null;

                        // First try exact InReplyTo match, fall back to ConversationId
                        if (!sentMailIndex.TryGetValue(inReplyToValue, out matchedSent))
                        {
                            string conversationId = inboxMail.ConversationId ?? string.Empty;
                            if (!conversationIndex.TryGetValue(conversationId, out matchedSent))
                                continue;
                        }

                        string replierAddress = inboxMail.From?.EmailAddress?.Address ?? string.Empty;
                        if (string.IsNullOrEmpty(replierAddress)) continue;

                        bool isOriginalToRecipient = matchedSent.To.Any(x =>
                            x.Equals(replierAddress, StringComparison.OrdinalIgnoreCase));

                        if (!isOriginalToRecipient) continue;

                        matchedSent.RepliedByAddresses.Add(replierAddress);
                        matchedSent.NotRepliedAddresses = matchedSent.NotRepliedAddresses
                            .Where(addr => !addr.Equals(replierAddress, StringComparison.OrdinalIgnoreCase))
                            .ToList();
                        matchedSent.HasPendingReplies = matchedSent.NotRepliedAddresses.Count > 0;
                    }

                    skip += pageSize;
                }
                while (inboxPage.Count == pageSize);

                return mails;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error fetching sent mails: " + ex.Message);
                throw;
            }
        }

        // GRAPH VERSION: paginates via $skip (same mechanism GetInboxx already uses)
        // instead of EWS's manual Offset do-while loop. internetMessageId is a
        // first-class Graph field, but it turned out to be computed and never
        // actually used in the original method's DTO output, so it's dropped here
        // rather than carried over as dead code. In-Reply-To still has to come from
        // internetMessageHeaders -- Graph doesn't expose it as its own field any more
        // directly than EWS's own fallback path did.
        public async Task<List<SentMailTrackDto>> GetOneDayMails(string targetUserEmail)
        {
            try
            {
                DateTime startDate = DateTime.UtcNow.AddDays(-7);
                var startDateIso = startDate.ToString("yyyy-MM-ddTHH:mm:ssZ");

                var mails = new List<SentMailTrackDto>();
                const int pageSize = 500;
                int skip = 0;
                List<Message> pageItems;

                do
                {
                    var page = await _graphClient.Users[targetUserEmail].MailFolders["Inbox"].Messages
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.Filter = $"receivedDateTime ge {startDateIso}";
                            requestConfig.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                            requestConfig.QueryParameters.Select = new[]
                            {
                                "id", "subject", "receivedDateTime", "from", "isRead",
                                "conversationId", "toRecipients", "ccRecipients", "body", "internetMessageHeaders"
                            };
                            requestConfig.QueryParameters.Top = pageSize;
                            requestConfig.QueryParameters.Skip = skip;
                            requestConfig.Headers.Add("Prefer", "IdType=\"ImmutableId\"");
                        });

                    pageItems = page?.Value ?? new List<Message>();

                    foreach (var inboxMail in pageItems)
                    {
                        string inReplyToValue = inboxMail.InternetMessageHeaders?
                            .FirstOrDefault(h => string.Equals(h.Name, "In-Reply-To", StringComparison.OrdinalIgnoreCase))
                            ?.Value ?? string.Empty;

                        bool isReply = !string.IsNullOrWhiteSpace(inReplyToValue);
                        string emailBody = inboxMail.Body?.Content ?? string.Empty;
                        bool hasQuestionMark = ContainsQuestionMark(emailBody);

                        mails.Add(new SentMailTrackDto
                        {
                            MessageId = inboxMail.Id ?? string.Empty,
                            ConversationId = inboxMail.ConversationId ?? string.Empty,
                            To = (inboxMail.ToRecipients ?? new List<Recipient>())
                                .Select(r => r.EmailAddress?.Address ?? string.Empty).ToList(),
                            Cc = (inboxMail.CcRecipients ?? new List<Recipient>())
                                .Select(r => r.EmailAddress?.Address ?? string.Empty).ToList(),
                            SentAt = inboxMail.ReceivedDateTime?.UtcDateTime ?? default,
                            Subject = inboxMail.Subject ?? string.Empty,
                            IsReply = isReply,
                            RepliedByAddresses = new HashSet<string>(),
                            NotRepliedAddresses = new List<string>(),
                            HasPendingReplies = false,
                            HasQuestionMark = hasQuestionMark,
                            Body = emailBody,
                            From = inboxMail.From?.EmailAddress?.Address ?? string.Empty,
                            IsRead = inboxMail.IsRead ?? false
                        });
                    }

                    skip += pageSize;
                }
                while (pageItems.Count == pageSize);

                return mails;
            }
            catch (Exception ex)
            {
                // log ex if needed
                Console.WriteLine("Error fetching one-day mails: " + ex.Message);
                return new List<SentMailTrackDto>();
            }
        }

        // GRAPH VERSION: fetches the target message once (to learn its conversationId),
        // then ONE query for every message sharing that conversationId. Since
        // /users/{mbx}/messages with no folder scope searches the WHOLE mailbox by
        // default, this actually improves on the EWS version, which only checked
        // Inbox + Sent Items and could miss thread messages filed elsewhere (a
        // subfolder, Drafts, etc.) -- matches your team lead's note that
        // conversationId is now a first-class field.
        // Attachment content is NOT loaded here (metadata only, same as EWS did) --
        // call DownloadAttachment for actual bytes once a specific one is opened.
        // (Also dropped a dead local -- the original built a "mail" object from the
        // single bound email that was never actually used; only the thread loop's
        // "mails" list was returned.)
        public async Task<List<MailDto>> GetOneMail(string messageId, string targetUserEmail)
        {
            try
            {
                var selectFields = new[]
                {
                    "id", "subject", "body", "receivedDateTime", "from",
                    "toRecipients", "ccRecipients", "conversationId", "hasAttachments"
                };

                var target = await _graphClient.Users[targetUserEmail].Messages[messageId]
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Select = selectFields;
                        requestConfig.Headers.Add("Prefer", "IdType=\"ImmutableId\"");
                    });

                var conversationId = target?.ConversationId ?? string.Empty;

                var threadResult = await _graphClient.Users[targetUserEmail].Messages
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Filter = $"conversationId eq '{conversationId}'";
                        requestConfig.QueryParameters.Select = selectFields;
                        requestConfig.QueryParameters.Top = 100;
                        requestConfig.Headers.Add("Prefer", "IdType=\"ImmutableId\"");
                    });

                var mails = new List<MailDto>();

                foreach (var item in threadResult?.Value ?? new List<Message>())
                {
                    var attachments = new List<HelpdeskAttachment>();

                    if (item.HasAttachments == true)
                    {
                        var attResult = await _graphClient.Users[targetUserEmail]
                            .Messages[item.Id]
                            .Attachments
                            .GetAsync(requestConfig =>
                            {
                                requestConfig.QueryParameters.Select = new[] { "id", "name", "contentType", "size" };
                            });

                        foreach (var att in (attResult?.Value ?? new List<GraphAttachment>()).OfType<FileAttachment>())
                        {
                            attachments.Add(new HelpdeskAttachment
                            {
                                Name = att.Name,
                                Type = att.ContentType,
                                Size = att.Size ?? 0,
                                AttachmentId = att.Id
                            });
                        }
                    }

                    mails.Add(new MailDto
                    {
                        From = item.From?.EmailAddress?.Address ?? string.Empty,
                        Subject = item.Subject ?? string.Empty,
                        Body = item.Body?.Content ?? string.Empty,
                        Received = item.ReceivedDateTime?.UtcDateTime ?? default,
                        MessageId = item.Id ?? string.Empty,
                        To = (item.ToRecipients ?? new List<Recipient>())
                            .Select(x => x.EmailAddress?.Address ?? string.Empty).ToList(),
                        Cc = (item.CcRecipients ?? new List<Recipient>())
                            .Select(x => x.EmailAddress?.Address ?? string.Empty).ToList(),
                        Attachments = attachments,
                        ConversationId = item.ConversationId ?? string.Empty
                    });
                }

                return mails.OrderBy(x => x.Received).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error fetching mail: " + ex.Message);
                throw;
            }
        }

        // GRAPH VERSION: no single Graph call replicates this exactly. UnreadCount
        // here still means "how many unread items came back in this fetch," capped
        // at `count` -- same as the original (it's not the mailbox's true total
        // unread count; if you actually want that, Users[mbx].MailFolders["Inbox"]
        // exposes a free UnreadItemCount field with no message fetch required at all
        // -- swap to that if this method's current behavior turns out to be a bug
        // rather than an intentional cap).
        // Recipients come straight off the same $select'd fetch (Graph doesn't need
        // EWS's two-phase find-then-load-properties dance). Attachment file types
        // still need one Attachments-list call per message that has them, same N+1
        // shape as GetOneMail, kept lightweight via $select to skip content bytes.
        public async Task<MailStatsDto> GetMailStats(string targetUserEmail, int count)
        {
            try
            {
                var result = await _graphClient.Users[targetUserEmail].MailFolders["Inbox"].Messages
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Filter = "isRead eq false";
                        requestConfig.QueryParameters.Select = new[]
                        {
                            "id", "toRecipients", "ccRecipients", "hasAttachments"
                        };
                        requestConfig.QueryParameters.Top = count;
                    });

                var items = result?.Value ?? new List<Message>();
                int unreadCount = items.Count;

                var uniqueToRecipients = items
                    .SelectMany(m => m.ToRecipients ?? new List<Recipient>())
                    .Select(r => r.EmailAddress?.Address?.ToLower() ?? string.Empty)
                    .Where(a => !string.IsNullOrEmpty(a))
                    .Distinct()
                    .Count();

                var uniqueCCRecipients = items
                    .SelectMany(m => m.CcRecipients ?? new List<Recipient>())
                    .Select(r => r.EmailAddress?.Address?.ToLower() ?? string.Empty)
                    .Where(a => !string.IsNullOrEmpty(a))
                    .Distinct()
                    .Count();

                var itemsWithAttachments = items.Where(i => i.HasAttachments == true).ToList();

                var fileTypeCounts = new Dictionary<string, int>();
                int totalAttachments = 0;

                foreach (var item in itemsWithAttachments)
                {
                    var attResult = await _graphClient.Users[targetUserEmail]
                        .Messages[item.Id]
                        .Attachments
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.Select = new[] { "name" };
                        });

                    foreach (var att in attResult?.Value ?? new List<GraphAttachment>())
                    {
                        totalAttachments++;
                        string ext = Path.GetExtension(att.Name)?.ToLower();
                        if (string.IsNullOrEmpty(ext)) ext = "unknown";

                        if (fileTypeCounts.ContainsKey(ext))
                            fileTypeCounts[ext]++;
                        else
                            fileTypeCounts[ext] = 1;
                    }
                }

                return new MailStatsDto
                {
                    UnreadCount = unreadCount,
                    TotalAttachments = totalAttachments,
                    TotalToRecipients = uniqueToRecipients,
                    TotalCcRecipients = uniqueCCRecipients,
                    FileTypeCounts = fileTypeCounts
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error fetching mail stats: " + ex.Message);
                throw;
            }
        }
        // GRAPH VERSION: since the caller already knows the specific attachmentId,
        // fetch that one directly instead of loading every attachment on the message
        // and filtering client-side like the EWS version did.
        public async Task<(byte[] Content, string FileName, string ContentType)> DownloadAttachment(string messageId, string attachmentId, string targetUserEmail)
        {
            try
            {
                var attachment = await _graphClient.Users[targetUserEmail]
                    .Messages[messageId]
                    .Attachments[attachmentId]
                    .GetAsync();

                if (attachment is not FileAttachment fileAtt)
                {
                    throw new Exception("Attachment not found");
                }

                return (fileAtt.ContentBytes ?? Array.Empty<byte>(), fileAtt.Name ?? string.Empty, fileAtt.ContentType ?? string.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error downloading attachment: " + ex.Message);
                throw new Exception("Attachment not found");
            }
        }


        // GRAPH VERSION: EWS did Bind -> set IsRead -> Update (3 round trips' worth of
        // ceremony). Graph collapses this to one PATCH with just the field that changed.
        public async Task<bool> SetMailReadStatus(
            string targetUserEmail,
            string messageId,
            bool isRead
        )
        {
            try
            {
                await _graphClient.Users[targetUserEmail].Messages[messageId]
                    .PatchAsync(new Message { IsRead = isRead });

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }


        // list [] not equal to null

        // GRAPH VERSION: attachments are inlined as base64 contentBytes in the same
        // one-shot sendMail call -- works for Graph's documented ~3MB per-attachment
        // limit. NOT yet handled: attachments over that size, which need Graph's
        // separate upload-session flow (create a draft, PUT the file in byte-range
        // chunks, then send the draft) -- a materially bigger addition than the rest
        // of this method, so flagging rather than guessing at it here.
        public async Task<bool> SendMailAsServiceAccount(
            List<string> to,
            List<string>? cc,
            string subject,
            string bodyHtml,
            List<IFormFile>? attachments,
            string from
        )
        {
            try
            {
                if (string.IsNullOrWhiteSpace(from))
                    throw new ArgumentException("From address is required.");

                if ((to == null || !to.Any()) && (cc == null || !cc.Any()))
                    throw new ArgumentException("At least one TO or CC recipient is required.");

                var message = new Message
                {
                    Subject = subject,
                    Body = new ItemBody { ContentType = BodyType.Html, Content = bodyHtml },
                    ToRecipients = (to ?? new List<string>())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(t => new Recipient { EmailAddress = new Microsoft.Graph.Models.EmailAddress { Address = t.Trim() } })
                        .ToList(),
                    CcRecipients = (cc ?? new List<string>())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(c => new Recipient { EmailAddress = new Microsoft.Graph.Models.EmailAddress { Address = c.Trim() } })
                        .ToList(),
                    Attachments = new List<GraphAttachment>()
                };

                if (attachments != null)
                {
                    foreach (var file in attachments.Where(f => f.Length > 0))
                    {
                        using var ms = new MemoryStream();
                        file.CopyTo(ms);

                        message.Attachments.Add(new FileAttachment
                        {
                            OdataType = "#microsoft.graph.fileAttachment",
                            Name = file.FileName,
                            ContentType = file.ContentType,
                            ContentBytes = ms.ToArray()
                        });
                    }
                }

                if (!message.ToRecipients.Any() && !message.CcRecipients.Any())
                {
                    throw new Exception("No valid recipients found.");
                }

                await _graphClient.Users[from].SendMail.PostAsync(new SendMailPostRequestBody
                {
                    Message = message,
                    SaveToSentItems = true
                });

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Email Error: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
                return false;
            }
        }


        //       public bool ReplyToMail(
        //           string messageItemId,
        //           string replyBodyHtml,
        //           List<string> toList,
        //           List<string>? ccList,
        //           List<IFormFile>? attachments,
        //           string targetUserEmail,
        //           bool isReplyAll
        //       )
        //       {
        //           try
        //           {
        //               ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        //               var service = new ExchangeService
        //               {
        //                   Credentials = new WebCredentials(
        //                       "hfmgemailservice@hfmg.net",
        //                       "Top99secret!"
        //                   ),
        //                   Url = new Uri("https://mail.hfmg.net/EWS/Exchange.asmx")
        //               };

        //               //impersonate
        //               service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, targetUserEmail);

        //               // 1️⃣ Bind original email
        //               EmailMessage original = EmailMessage.Bind(
        //    service,
        //    new ItemId(messageItemId),
        //    new PropertySet(BasePropertySet.FirstClassProperties)
        //);

        //               // 2️⃣ Create reply
        //               ResponseMessage response = original.CreateReply(false);

        //               response.ToRecipients.Clear();
        //               response.CcRecipients.Clear();

        //               foreach (var to in toList)
        //                   response.ToRecipients.Add(to);

        //               if (ccList != null && ccList.Any())
        //                   foreach (var cc in ccList)
        //                       response.CcRecipients.Add(cc);

        //               response.BodyPrefix = new MessageBody(
        //                   EwsBodyType.HTML,
        //                   replyBodyHtml
        //               );

        //               // 3️⃣ SAVE response as EmailMessage
        //               EmailMessage replyMessage = response.Save();

        //               // 4️⃣ Add attachments to EmailMessage
        //               if (attachments != null && attachments.Any())
        //               {
        //                   foreach (var file in attachments)
        //                   {
        //                       using var ms = new MemoryStream();
        //                       file.CopyTo(ms);

        //                       replyMessage.Attachments.AddFileAttachment(
        //                           file.FileName,
        //                           ms.ToArray()
        //                       );
        //                   }

        //                   // MUST call Update after adding attachments
        //                   replyMessage.Update(ConflictResolutionMode.AlwaysOverwrite);
        //               }

        //               // 5️⃣ Send
        //               replyMessage.SendAndSaveCopy();
        //               return true;
        //           }
        //           catch (Exception ex)
        //           {
        //               // log ex if needed
        //               Console.Write($"this is the error{ex}");
        //               return false;
        //           }
        //       }



        //send mail for tracking
        // GRAPH VERSION: same sendMail pattern as SendMailAsServiceAccount, sent from
        // _reminderMailbox (Graph:ReminderMailbox in config) instead of a hardcoded
        // literal. Same attachment-size caveat applies here too.
        public async Task<bool> SendMailTracking(
            List<string> to,
            List<string>? cc,
            string subject,
            string bodyHtml,
            List<IFormFile>? attachments
        )
        {
            try
            {
                if (to == null || !to.Any())
                    throw new ArgumentException("At least one TO recipient is required.");

                var message = new Message
                {
                    Subject = subject,
                    Body = new ItemBody { ContentType = BodyType.Html, Content = bodyHtml },
                    ToRecipients = to
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(t => new Recipient { EmailAddress = new EmailAddress { Address = t.Trim() } })
                        .ToList(),
                    CcRecipients = (cc ?? new List<string>())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(c => new Recipient { EmailAddress = new EmailAddress { Address = c.Trim() } })
                        .ToList(),
                    Attachments = new List<GraphAttachment>()
                };

                if (attachments != null)
                {
                    foreach (var file in attachments.Where(f => f.Length > 0))
                    {
                        using var ms = new MemoryStream();
                        file.CopyTo(ms);

                        message.Attachments.Add(new FileAttachment
                        {
                            OdataType = "#microsoft.graph.fileAttachment",
                            Name = file.FileName,
                            ContentType = file.ContentType,
                            ContentBytes = ms.ToArray()
                        });
                    }
                }

                if (!message.ToRecipients.Any() && !message.CcRecipients.Any())
                {
                    throw new Exception("No valid recipients found.");
                }

                await _graphClient.Users[_reminderMailbox].SendMail.PostAsync(new SendMailPostRequestBody
                {
                    Message = message,
                    SaveToSentItems = true
                });
                Console.WriteLine(_reminderMailbox);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Email Error: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
                throw;
           
            }
        }




        //reply and replyall




        // GRAPH VERSION: same create-draft -> modify -> attach -> send shape as the
        // EWS version (CreateReply/CreateReplyAll -> Save -> attach -> Update -> Send).
        // Reply/replyAll's one-shot request body CAN set recipients directly, but
        // since it's unclear whether that ADDS to or REPLACES the auto-computed
        // reply-all recipient set, this creates an empty draft first and applies the
        // same clear-or-append logic the original EWS code used, so the reply-all
        // "don't drop existing recipients, just add extras" behavior is guaranteed
        // rather than assumed.
        // Body handling: EWS's BodyPrefix PREPENDS your text above the quoted original
        // without touching the quote chain below it. Graph only has one Body field, and
        // CreateReply's draft already comes back with the quote chain baked in, so a
        // straight overwrite would erase it. This reads the draft's existing Body.Content
        // and prepends the new text to it instead, to get the same "your reply above,
        // original message quoted below" result.
        public async Task<bool> ReplyToMail(
            string messageItemId,
            string replyBodyHtml,
            List<string>? toList,
            List<string>? ccList,
            List<IFormFile>? attachments,
            string targetUserEmail,
            bool isReplyAll
        )
        {
            try
            {
                Message draft = isReplyAll
     ? await _graphClient.Users[targetUserEmail].Messages[messageItemId]
           .CreateReplyAll.PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.CreateReplyAll.CreateReplyAllPostRequestBody())
     : await _graphClient.Users[targetUserEmail].Messages[messageItemId]
           .CreateReply.PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.CreateReply.CreateReplyPostRequestBody());

                var draftId = draft.Id;
                var newToRecipients = new List<Recipient>(draft.ToRecipients ?? new List<Recipient>());
                var newCcRecipients = new List<Recipient>(draft.CcRecipients ?? new List<Recipient>());

                if (!isReplyAll)
                {
                    newToRecipients.Clear();
                    newCcRecipients.Clear();

                    foreach (var to in (toList ?? new List<string>()).Where(x => !string.IsNullOrWhiteSpace(x)))
                        newToRecipients.Add(new Recipient { EmailAddress = new EmailAddress { Address = to.Trim() } });

                    foreach (var cc in (ccList ?? new List<string>()).Where(x => !string.IsNullOrWhiteSpace(x)))
                        newCcRecipients.Add(new Recipient { EmailAddress = new EmailAddress { Address = cc.Trim() } });
                }
                else
                {
                    // REPLY ALL -> keep existing recipients, only add extras not already present
                    foreach (var to in (toList ?? new List<string>()).Where(x => !string.IsNullOrWhiteSpace(x)))
                    {
                        if (!newToRecipients.Any(r => string.Equals(r.EmailAddress?.Address, to, StringComparison.OrdinalIgnoreCase)))
                            newToRecipients.Add(new Recipient { EmailAddress = new EmailAddress { Address = to.Trim() } });
                    }

                    foreach (var cc in (ccList ?? new List<string>()).Where(x => !string.IsNullOrWhiteSpace(x)))
                    {
                        if (!newCcRecipients.Any(r => string.Equals(r.EmailAddress?.Address, cc, StringComparison.OrdinalIgnoreCase)))
                            newCcRecipients.Add(new Recipient { EmailAddress = new EmailAddress { Address = cc.Trim() } });
                    }
                }

                var combinedBody = replyBodyHtml + (draft.Body?.Content ?? string.Empty);

                await _graphClient.Users[targetUserEmail].Messages[draftId].PatchAsync(new Message
                {
                    ToRecipients = newToRecipients,
                    CcRecipients = newCcRecipients,
                    Body = new ItemBody { ContentType = BodyType.Html, Content = combinedBody }
                });

                if (attachments != null)
                {
                    foreach (var file in attachments.Where(f => f.Length > 0))
                    {
                        using var ms = new MemoryStream();
                        file.CopyTo(ms);

                        await _graphClient.Users[targetUserEmail].Messages[draftId].Attachments.PostAsync(new FileAttachment
                        {
                            OdataType = "#microsoft.graph.fileAttachment",
                            Name = file.FileName,
                            ContentType = file.ContentType,
                            ContentBytes = ms.ToArray()
                        });
                    }
                }

                await _graphClient.Users[targetUserEmail].Messages[draftId].Send.PostAsync();

                return true;
            }
            catch (Exception ex)
            {
                if (ex is ODataError odataEx)
                {
                    Console.WriteLine($"Graph Error Code: {odataEx.Error?.Code}, Message: {odataEx.Error?.Message}");
                    Console.WriteLine($"Inner: {odataEx.InnerException}");
                }
                else
                {
                    Console.WriteLine($"Reply Error: {ex}"); // ex.ToString(), not ex.Message — get the stack trace
                }
                return false;
            }
        }

        //forwarding email

        // GRAPH VERSION: same create-draft -> modify -> attach -> send shape used for
        // ReplyToMail, including the same BodyPrefix-preservation approach (prepend to
        // the draft's already-quoted body instead of overwriting it).
        // Kept the original's explicit "also copy the original message's attachments
        // onto the forward draft" step even though createForward likely includes them
        // by default -- matching the original code's evident intent rather than
        // assuming that behavior away.
        public async Task<(bool success, string error)> ForwordEmail(ForwardMailDto details, string targetUserEmail)
        {
            try
            {
                Message draft = await _graphClient.Users[targetUserEmail]
    .Messages[details.MessageItemId]
    .CreateForward.PostAsync(new CreateForwardPostRequestBody());

                var draftId = draft.Id;

                if (string.IsNullOrEmpty(draftId))
                    return (false, "CreateForward did not return a valid draft id");

                var newToRecipients = (details.ToList ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(to => new Recipient { EmailAddress = new EmailAddress { Address = to.Trim() } })
                    .ToList();

                var newCcRecipients = (details.CcList ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(cc => new Recipient { EmailAddress = new EmailAddress { Address = cc.Trim() } })
                    .ToList();

                var combinedBody = details.Body + (draft.Body?.Content ?? string.Empty);

                await _graphClient.Users[targetUserEmail].Messages[draftId].PatchAsync(new Message
                {
                    ToRecipients = newToRecipients,
                    CcRecipients = newCcRecipients,
                    Body = new ItemBody { ContentType = BodyType.Html, Content = combinedBody }
                });

                if (details.Attachments != null)
                {
                    foreach (var file in details.Attachments.Where(f => f.Length > 0))
                    {
                        using var ms = new MemoryStream();
                        file.CopyTo(ms);

                        await _graphClient.Users[targetUserEmail].Messages[draftId].Attachments.PostAsync(new FileAttachment
                        {
                            OdataType = "#microsoft.graph.fileAttachment",
                            Name = file.FileName,
                            ContentType = file.ContentType,
                            ContentBytes = ms.ToArray()
                        });
                    }
                }

                // Explicitly copy the original message's attachments too (see note above)
                var originalAttachments = await _graphClient.Users[targetUserEmail]
                    .Messages[details.MessageItemId]
                    .Attachments
                    .GetAsync();

                foreach (var att in (originalAttachments?.Value ?? new List<GraphAttachment>()).OfType<FileAttachment>())
                {
                    await _graphClient.Users[targetUserEmail].Messages[draftId].Attachments.PostAsync(new FileAttachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        Name = att.Name,
                        ContentType = att.ContentType,
                        ContentBytes = att.ContentBytes
                    });
                }

                await _graphClient.Users[targetUserEmail].Messages[draftId].Send.PostAsync();

                return (true, null);
            }
            catch (Exception ex)
            {
                string errorMessage;

                if (ex is ODataError odataEx)
                {
                    errorMessage = $"Graph Error Code: {odataEx.Error?.Code}, Message: {odataEx.Error?.Message}";
                    Console.WriteLine(errorMessage);
                    Console.WriteLine($"Inner: {odataEx.InnerException}");
                }
                else
                {
                    errorMessage = ex.Message;
                    Console.WriteLine($"Forward Error: {ex}"); // full ToString() in console for stack trace
                }

                return (false, errorMessage);
            }
        }


        //get mail by mailid

        //public MailDto GetOneMail(string messageId, string targetUserEmail)
        //{
        //    try
        //    {
        //        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        //        var service = new ExchangeService
        //        {
        //            Credentials = new WebCredentials(
        //                "hfmgemailservice@hfmg.net",
        //                "Top99secret!"
        //            ),
        //            Url = new Uri("https://mail.hfmg.net/EWS/Exchange.asmx")
        //        };

        //        // ✅ Impersonation
        //        service.ImpersonatedUserId = new ImpersonatedUserId(
        //            ConnectingIdType.SmtpAddress,
        //            targetUserEmail
        //        );

        //        // ✅ Bind email with full properties
        //        EmailMessage email = EmailMessage.Bind(
        //            service,
        //            new ItemId(messageId),
        //            new PropertySet(
        //                BasePropertySet.FirstClassProperties,
        //                ItemSchema.Attachments
        //            )
        //        );

        //        // ✅ Convert HTML → Plain text (your existing method)
        //        string plainBody = GetPlainTextFromHtml(email.Body?.Text);

        //        var mail = new MailDto
        //        {
        //            From = email.From?.Address ?? string.Empty,
        //            Subject = email.Subject ?? string.Empty,
        //            Body = plainBody ?? string.Empty,
        //            Received = email.DateTimeReceived,
        //            MessageId = email.Id?.UniqueId ?? string.Empty,
        //            To = email.ToRecipients.Select(r => r.Address).ToList(),
        //            Cc = email.CcRecipients.Select(r => r.Address).ToList()
        //        };

        //        // ✅ Attachments
        //        if (email.HasAttachments)
        //        {
        //            foreach (var att in email.Attachments)
        //            {
        //                if (att is FileAttachment fileAtt)
        //                {
        //                    fileAtt.Load();

        //                    mail.Attachments.Add(new HelpdeskAttachment
        //                    {
        //                        Name = fileAtt.Name,
        //                        Type = fileAtt.ContentType,
        //                        //Hash = Convert.ToBase64String(fileAtt.Content)
        //                        //Size = attachment.Size,
        //                        //AttachmentId = attachment.Id
        //                    });
        //                }
        //            }
        //        }

        //        return mail;







        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error fetching mail: " + ex.Message);
        //        throw; // ✅ Let controller handle
        //    }
        //}



        // GRAPH VERSION: calendarView takes the date range as query params instead of
        // a CalendarView object. EWS's single "Appointment" concept splits here into
        // Event + a flat Attendees list -- Required/Optional now come from filtering
        // Attendee.Type instead of two separate recipient collections.
        public async Task<List<CalendarDto>> GetCalendarMeetings(string targetUserEmail, DateTime startDate, DateTime endDate)
        {
            var meetings = new List<CalendarDto>();

            try
            {
                Console.WriteLine($"Received startDate={startDate:O}, endDate={endDate:O}");

                var calendars = await _graphClient.Users[targetUserEmail].Calendars.GetAsync();

                foreach (var calendar in calendars?.Value ?? new List<Calendar>())
                {
                    var allEvents = new List<Event>();

                    var result = await _graphClient.Users[targetUserEmail].Calendars[calendar.Id].CalendarView
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.StartDateTime = startDate.ToString("o");
                            requestConfig.QueryParameters.EndDateTime = endDate.ToString("o");
                            requestConfig.QueryParameters.Select = new[]
                            {
                        "id", "subject", "body", "start", "end", "location",
                        "organizer", "attendees", "isAllDay", "isCancelled", "responseStatus"
                            };
                            requestConfig.QueryParameters.Top = 999;
                            requestConfig.Headers.Add("Prefer", "IdType=\"ImmutableId\"");
                        });

                    allEvents.AddRange(result?.Value ?? new List<Event>());

                    var nextLink = result?.OdataNextLink;
                    while (!string.IsNullOrEmpty(nextLink))
                    {
                        var nextPageBuilder = new Microsoft.Graph.Users.Item.Calendars.Item.CalendarView.CalendarViewRequestBuilder(
                            nextLink, _graphClient.RequestAdapter);
                        var nextPage = await nextPageBuilder.GetAsync();
                        allEvents.AddRange(nextPage?.Value ?? new List<Event>());
                        nextLink = nextPage?.OdataNextLink;
                    }

                    foreach (var appointment in allEvents)
                    {
                        var allAttendees = appointment.Attendees ?? new List<Attendee>();
                        meetings.Add(new CalendarDto
                        {
                            Subject = appointment.Subject ?? string.Empty,
                            Body = GetPlainTextFromHtml(appointment.Body?.Content ?? string.Empty),
                            Start = ParseGraphDateTime(appointment.Start),
                            End = ParseGraphDateTime(appointment.End),
                            Location = appointment.Location?.DisplayName ?? string.Empty,
                            Organizer = appointment.Organizer?.EmailAddress?.Address ?? string.Empty,
                            RequiredAttendees = allAttendees.Where(a => a.Type == AttendeeType.Required)
                                .Select(a => a.EmailAddress?.Address ?? string.Empty).ToList(),
                            OptionalAttendees = allAttendees.Where(a => a.Type == AttendeeType.Optional)
                                .Select(a => a.EmailAddress?.Address ?? string.Empty).ToList(),
                            IsAllDay = appointment.IsAllDay ?? false,
                            IsCancelled = appointment.IsCancelled ?? false,
                            MeetingId = appointment.Id ?? string.Empty,
                            ResponseStatus = appointment.ResponseStatus?.Response?.ToString() ?? string.Empty
                        });
                    }
                }
            }
            catch (ODataError odErr)
            {
                throw new ApplicationException("Error communicating with Graph while reading calendar.", odErr);
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Unexpected error while reading calendar.", ex);
            }

            return meetings.OrderBy(m => m.Start).ToList();
        }

        private DateTime ParseGraphDateTime(DateTimeTimeZone? dtz)
        {
            if (dtz == null || string.IsNullOrEmpty(dtz.DateTime))
                return default;

            // Graph's DateTime string has no offset, e.g. "2026-07-13T09:00:00.0000000"
            var naive = DateTime.Parse(dtz.DateTime, null, System.Globalization.DateTimeStyles.RoundtripKind);

            if (string.IsNullOrEmpty(dtz.TimeZone) || dtz.TimeZone == "UTC")
                return DateTime.SpecifyKind(naive, DateTimeKind.Utc);

            try
            {
                var tzInfo = TZConvert.GetTimeZoneInfo(dtz.TimeZone); // via TimeZoneConverter NuGet package
                var utc = TimeZoneInfo.ConvertTimeToUtc(DateTime.SpecifyKind(naive, DateTimeKind.Unspecified), tzInfo);
                return utc;
            }
            catch (TimeZoneNotFoundException)
            {
                return DateTime.SpecifyKind(naive, DateTimeKind.Unspecified);
            }
        }


        // GRAPH VERSION: EWS's appointment.Accept()/Decline() returned the updated
        // object directly. Graph's accept/decline are pure actions with no response
        // body, so this re-fetches the event afterward to build the returned DTO --
        // same idea as the original code's "reload appointment" step.
        // (Trimmed the original's verbose before/after Console.WriteLine debug dump --
        // it didn't affect the returned result, just noise. Shout if you want it back.)
        public async Task<CalendarDto> RespondToMeeting(string meetingId, MeetingAction action, string targetUserEmail)
        {
            try
            {
                var before = await _graphClient.Users[targetUserEmail].Events[meetingId]
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Select = new[] { "isCancelled" };
                    });

                if (before?.IsCancelled == true)
                {
                    throw new Exception("Meeting already cancelled.");
                }

                if (action == MeetingAction.Accept)
                {
                    await _graphClient.Users[targetUserEmail].Events[meetingId]
                        .Accept.PostAsync(new AcceptPostRequestBody { SendResponse = true });
                }
                else if (action == MeetingAction.Reject)
                {
                    await _graphClient.Users[targetUserEmail].Events[meetingId]
                        .Decline.PostAsync(new DeclinePostRequestBody { SendResponse = true });
                }
                else
                {
                    throw new Exception("Invalid action.");
                }

                var after = await _graphClient.Users[targetUserEmail].Events[meetingId]
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Select = new[]
                        {
                            "id", "subject", "start", "end", "isCancelled", "responseStatus"
                        };
                    });

                return new CalendarDto
                {
                    MeetingId = after?.Id ?? string.Empty,
                    Subject = after?.Subject ?? string.Empty,
                    Start = ParseGraphDateTime(after?.Start),
                    End = ParseGraphDateTime(after?.End),
                    IsCancelled = after?.IsCancelled ?? false,
                    ResponseStatus = after?.ResponseStatus?.Response?.ToString() ?? string.Empty
                };
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message, ex);
            }
        }


        //// GRAPH VERSION: targets the well-known "deleteditems" folder specifically. A
        //// message already in Deleted Items has nowhere left to be soft-deleted TO, so
        //// calling delete on it here purges it permanently — this is the "empty trash for
        //// one item" action, not reversible like the regular DeleteMail above.
        //public async Task<bool> DeleteMailFromDeletedItems(string messageId, string targetUserEmail)
        //{
        //    try
        //    {
        //        await _graphClient.Users[targetUserEmail]
        //            .MailFolders["deleteditems"]
        //            .Messages[messageId]
        //            .DeleteAsync();

        //        return true;
        //    }
        //    catch (ODataError odErr) when (odErr.ResponseStatusCode == 404)
        //    {
        //        // Already purged / never existed there — treat as success, nothing left to do
        //        return true;
        //    }
        //    catch (ODataError odErr)
        //    {
        //        Console.WriteLine($"Graph error purging mail {messageId} for {targetUserEmail}: " +
        //                           $"{odErr.Error?.Code} - {odErr.Error?.Message}");
        //        return false;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error purging mail: " + ex.Message);
        //        return false;
        //    }
        //}

        public List<MeetingInviteDto> GetMeetingRequests(string targetUserEmail, int count)
        {
            var invites = new List<MeetingInviteDto>();

            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var service = new ExchangeService
                {
                    Credentials = new WebCredentials("hfmgemailservice@hfmg.net", "Top99secret!"),
                    Url = new Uri("https://mail.hfmg.net/EWS/Exchange.asmx")
                };

                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, targetUserEmail);

                var mailbox = new EwsMailbox(targetUserEmail);
                var folderId = new FolderId(WellKnownFolderName.Inbox, mailbox);

                // 🔥 Filter only meeting requests
                SearchFilter filter = new SearchFilter.IsEqualTo(
                    ItemSchema.ItemClass,
                    "IPM.Schedule.Meeting.Request"
                );

                ItemView view = new ItemView(count)
                {
                    PropertySet = new PropertySet(BasePropertySet.FirstClassProperties)
                };

                FindItemsResults<Item> results =
                    service.FindItems(folderId, filter, view);

                foreach (Item item in results.Items)
                {
                    if (item is MeetingRequest meetingRequest)
                    {
                        meetingRequest.Load();

                        var invite = new MeetingInviteDto
                        {
                            MessageId = meetingRequest.Id?.UniqueId ?? string.Empty,
                            Subject = meetingRequest.Subject ?? string.Empty,
                            Body = GetPlainTextFromHtml(meetingRequest.Body?.Text ?? string.Empty),
                            Start = meetingRequest.Start,
                            End = meetingRequest.End,
                            Location = meetingRequest.Location ?? string.Empty,
                            Organizer = meetingRequest.Organizer?.Address ?? string.Empty,
                            RequiredAttendees = meetingRequest.RequiredAttendees
                                                        .Select(a => a.Address)
                                                        .ToList(),
                            OptionalAttendees = meetingRequest.OptionalAttendees
                                                        .Select(a => a.Address)
                                                        .ToList(),
                            IsRead = meetingRequest.IsRead
                        };

                        invites.Add(invite);
                    }
                }
            }
            catch (ServiceRequestException ex)
            {
                throw new ApplicationException("Error communicating with Exchange server.", ex);
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Unexpected error while reading meeting requests.", ex);
            }

            return invites;
        }





        public class MeetingInviteDto
        {
            public string MessageId { get; set; }
            public string Subject { get; set; }
            public string Body { get; set; }
            public DateTime Start { get; set; }
            public DateTime End { get; set; }
            public string Location { get; set; }
            public string Organizer { get; set; }
            public List<string> RequiredAttendees { get; set; } = new();
            public List<string> OptionalAttendees { get; set; } = new();
            public bool IsRead { get; set; }
        }


        public async Task<PersonalContactsResult> SearchPersonalContacts(string targetUserEmail, string query)
        {
            var matches = new List<PersonalContactDto>();
            var queryLower = query.ToLowerInvariant();

            try
            {
                var defaultResult = await _graphClient.Users[targetUserEmail].Contacts
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Select = new[] { "displayName", "emailAddresses" };
                        requestConfig.QueryParameters.Top = 999;
                    });

                AddMatches(defaultResult?.Value, queryLower, matches);

                var foldersResult = await _graphClient.Users[targetUserEmail].ContactFolders
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Select = new[] { "id", "displayName" };
                    });

                Console.WriteLine($"ContactFolders found: {foldersResult?.Value?.Count ?? 0}");
                foreach (var f in foldersResult?.Value ?? new List<ContactFolder>())
                    Console.WriteLine($"  Folder: {f.DisplayName} ({f.Id})");

                foreach (var folder in foldersResult?.Value ?? new List<ContactFolder>())
                {
                    var folderContacts = await _graphClient.Users[targetUserEmail]
                        .ContactFolders[folder.Id].Contacts
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.Select = new[] { "displayName", "emailAddresses" };
                            requestConfig.QueryParameters.Top = 999;
                        });

                    AddMatches(folderContacts?.Value, queryLower, matches);
                }

                var deduped = matches
                    .GroupBy(m => m.Email.ToLowerInvariant())
                    .Select(g => g.First())
                    .ToList();

                return new PersonalContactsResult { Success = true, Contacts = deduped };
            }
            catch (ODataError odErr)
            {
                var code = odErr.Error?.Code ?? "Unknown";
                var msg = odErr.Error?.Message ?? odErr.Message;
                Console.WriteLine($"Graph error searching personal contacts: {code} - {msg}");
                return new PersonalContactsResult
                {
                    Success = false,
                    ErrorCode = code,
                    ErrorMessage = msg
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error searching personal contacts: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
                return new PersonalContactsResult
                {
                    Success = false,
                    ErrorCode = "Exception",
                    ErrorMessage = ex.Message
                };
            }
        }

        public class PersonalContactsResult
        {
            public bool Success { get; set; }
            public List<PersonalContactDto> Contacts { get; set; } = new();
            public string? ErrorCode { get; set; }
            public string? ErrorMessage { get; set; }
        }

        public class PersonalContactDto
        {
            public string Email { get; set; }
            public string Name { get; set; }
            public string Source { get; set; }
        }
        private void AddMatches(List<Contact>? contacts, string queryLower, List<PersonalContactDto> matches)
        {
            foreach (var contact in contacts ?? new List<Contact>())
            {
                foreach (var addr in contact.EmailAddresses ?? new List<Microsoft.Graph.Models.EmailAddress>())
                {
                    if (string.IsNullOrEmpty(addr.Address)) continue;

                    bool nameMatch = (contact.DisplayName ?? "").ToLowerInvariant().Contains(queryLower);
                    bool emailMatch = addr.Address.ToLowerInvariant().Contains(queryLower);

                    if (nameMatch || emailMatch)
                    {
                        matches.Add(new PersonalContactDto
                        {
                            Email = addr.Address,
                            Name = contact.DisplayName ?? addr.Address,
                            Source = "personal"
                        });
                    }
                }
            }
        }

        

    }
}