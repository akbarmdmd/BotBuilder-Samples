using System;

namespace Microsoft.BotBuilderSamples.Models
{
    public class CustomFormResponse
    {
        public string Title { get; set; }
        public string Attendees { get; set; }
        public string Time { get; set; }
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

    public class EmailAddress
    {
        public string Name { get; set; }
        public string Address { get; set; }
    }
}
