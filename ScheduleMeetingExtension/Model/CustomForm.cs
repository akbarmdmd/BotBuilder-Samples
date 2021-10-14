using Microsoft.Graph;
using System;

namespace Microsoft.BotBuilderSamples.Models
{
    public class CustomFormResponse
    {
        public string Title { get; set; }
        public string Organizer { get; set; }
        public string AttendeeNames { get; set; }
        public string AttendeeAddresses { get; set; }
        public DateTime StartTime { get; set; }
        public string Duration { get; set; }
    }

    public class TeamsMessagesData
    {
        public Message[] Value { get; set; }
    }

    public class Message
    {
        public string InternetMessageId { get; set; }
        public string Subject { get; set; }
        public string BodyPreview { get; set; }
        public From Sender { get; set; }
    }

    public class From
    {
        public EmailAddress EmailAddress { get; set; }
    }
}
