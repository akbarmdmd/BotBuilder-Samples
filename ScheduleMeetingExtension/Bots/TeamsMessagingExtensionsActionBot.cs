// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.BotBuilderSamples.Helpers;
using Microsoft.BotBuilderSamples.Models;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class TeamsMessagingExtensionsActionBot : TeamsActivityHandler
    {
        public readonly string baseUrl;

        public TeamsMessagingExtensionsActionBot(IConfiguration configuration) : base()
        {
            this.baseUrl = configuration["BaseUrl"];
        }

        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            switch (action.CommandId)
            {
                case "scheduleMeeting":
                    //TODO: call create meeting api and return adaptive card in message.
                    return ScheduleMeetingResponse(turnContext, action);
            }
            return new MessagingExtensionActionResponse();
        }

        private MessagingExtensionActionResponse ScheduleMeetingResponse(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            // The user has chosen to create a card by choosing the 'Web View' context menu command.
            var a = action.Data.ToString();
            CustomFormResponse data = JsonConvert.DeserializeObject<CustomFormResponse>(a);

            var attachments = new List<MessagingExtensionAttachment>();
            attachments.Add(new MessagingExtensionAttachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>()
                    {
                      new AdaptiveTextInput() { Id = "title", Value = data.Title, Label = "Title" },
                      new AdaptiveTextInput() { Id = "attendees", Value = string.Join("; ", data.Attendees), Label = "Attendees" },
                      new AdaptiveTextInput() { Id = "time", Value = data.Time, Label = "Time" },
                    },
                },
            });

            return new MessagingExtensionActionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    //Type = "botMessagePreview",
                    //ActivityPreview = MessageFactory.Attachment(new Attachment
                    //{
                    //    Content = new AdaptiveCard("1.0")
                    //    {
                    //        Body = new List<AdaptiveElement>()
                    //            {
                    //            new AdaptiveTextBlock() { Text = "FormField1 value was:", Size = AdaptiveTextSize.Large },
                    //            new AdaptiveTextBlock() { Text = "Akbar" }
                    //            }
                    //    },
                    //    ContentType = AdaptiveCard.ContentType
                    //}) as Activity
                    //Type = "message",
                    //Text = "Meeting created"
                    Type = "result",
                    Attachments = attachments,
                    AttachmentLayout = "list",
                }
            };
        }

        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            switch (action.CommandId)
            {
                case "scheduleMeeting":
                    return MeetingDetails(turnContext, action);
                default:
                    // we are handling two cases within try/catch block 
                    //if the bot is installed it will create adaptive card attachment and show card with input fields
                    string memberName;
                    try
                    {
                        // Check if your app is installed by fetching member information.
                        var member = await TeamsInfo.GetMemberAsync(turnContext, turnContext.Activity.From.Id, cancellationToken);
                        memberName = member.Name;
                    }
                    catch (ErrorResponseException ex)
                    {
                        if (ex.Body.Error.Code == "BotNotInConversationRoster")
                        {
                            return new MessagingExtensionActionResponse
                            {
                                Task = new TaskModuleContinueResponse
                                {
                                    Value = new TaskModuleTaskInfo
                                    {
                                        Card = GetAdaptiveCardAttachmentFromFile("justintimeinstallation.json"),
                                        Height = 200,
                                        Width = 400,
                                        Title = "Adaptive Card - App Installation",
                                    },
                                },
                            };
                        }
                        throw; // It's a different error.
                    }

                    return new MessagingExtensionActionResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = new TaskModuleTaskInfo
                            {
                                Card = GetAdaptiveCardAttachmentFromFile("adaptiveCard.json"),
                                Height = 200,
                                Width = 400,
                                Title = $"Welcome {memberName}",
                            },
                        },
                    };
            }
        }

        private MessagingExtensionActionResponse MeetingDetails(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            // @esha fetch required properties from context or action, then call graph apis to fetch converstation details like conversation title, conversation attendees
            var response = new MessagingExtensionActionResponse()
            {
                Task = new TaskModuleContinueResponse()
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Height = 600,
                        Width = 450,
                        Title = "Schedule meeting",
                        Card = GetAdaptiveCardAttachmentForMeeting("Hack 2021", new string[2] { "eshalath@microsoft.com", "yashna@microsoft.com"}),
                    },
                },
            };
            return response;
        }

        private static Attachment GetAdaptiveCardAttachmentFromFile(string fileName)
        {
            //Read the card json and create attachment.
            string[] paths = { ".", "Resources", fileName };
            var adaptiveCardJson = File.ReadAllText(Path.Combine(paths));

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(adaptiveCardJson),
            };
            return adaptiveCardAttachment;
        }

        private static Attachment GetAdaptiveCardAttachmentForMeeting(string title, string[] attendees)
        {
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>()
                    {
                      new AdaptiveTextInput() { Id = "title", Value = title, Label = "Title" },
                      new AdaptiveTextInput() { Id = "attendees", Value = string.Join("; ", attendees), Label = "Attendees" },
                      new AdaptiveTextInput() { Id = "time", Value = string.Join(" - ", "03:00 pm", "04:00 pm"), Label = "Time" },
                    },
                    Actions = new List<AdaptiveAction>()
                    {
                      new AdaptiveSubmitAction()
                      {
                        Type = AdaptiveSubmitAction.TypeName,
                        Title = "Book",
                      },
                    },
                },
            };
            return adaptiveCardAttachment;
        }
    }
}
