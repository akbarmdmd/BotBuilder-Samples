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
using System.Net.Http;

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
                    return ScheduleMeetingResponse(turnContext, action);
            }
            return new MessagingExtensionActionResponse();
        }

        private MessagingExtensionActionResponse ScheduleMeetingResponse(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            // The user has chosen to create a card by choosing the 'Web View' context menu command.
            var a = action.Data.ToString();
            CustomFormResponse data = JsonConvert.DeserializeObject<CustomFormResponse>(a);

            //TODO: call create meeting api and on success return adaptive card in message.

            var attachments = new List<MessagingExtensionAttachment>();
            attachments.Add(new MessagingExtensionAttachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>()
                    {
                      new AdaptiveTextBlock() { Text = $"Meeting created with below details - ", Color = AdaptiveTextColor.Accent },
                      new AdaptiveTextBlock() { Text = $"**Title** : {data.Title}" },
                      new AdaptiveTextBlock() { Text = $"**Attendees** : {string.Join("; ", data.Attendees)}" },
                      new AdaptiveTextBlock() { Text = $"**Time** : {data.Time}" },
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
                    return await MeetingDetails(turnContext, action);
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

        private async Task<MessagingExtensionActionResponse> MeetingDetails(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            // TODO Fetch required properties from context or action, then call graph apis to fetch converstation details like
            // conversation title, conversation attendees, time (using find meeting times or LU on message)
            var groupId = "b77dde63-1f70-4668-95d5-0c34e038514b";
            var tenantId = "72f988bf-86f1-41af-91ab-2d7cd011db47";
            var threadId = "19:dbad371212c0450f8c63da35ef7a0484@thread.tacv2";
            var conversationId = "19:dbad371212c0450f8c63da35ef7a0484@thread.tacv2;messageid=1634033527688";

            var token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IlJxc1dwRTkwN2gwZTBxejBEdHBSZG9hY3AyUVlocUtDWVFSNUdoNTRwZWMiLCJhbGciOiJSUzI1NiIsIng1dCI6Imwzc1EtNTBjQ0g0eEJWWkxIVEd3blNSNzY4MCIsImtpZCI6Imwzc1EtNTBjQ0g0eEJWWkxIVEd3blNSNzY4MCJ9.eyJhdWQiOiJodHRwczovL3N1YnN0cmF0ZS5vZmZpY2UuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3LyIsImlhdCI6MTYzNDEyMzI4MiwibmJmIjoxNjM0MTIzMjgyLCJleHAiOjE2MzQxMjcxODIsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVVFBdS84VEFBQUE4NnlBV3JhMGFwd1FQUGd6TzlUT1JmY0dJYktOTUhRUUxXcGozaC9qdmRNUFMvcG5KcExNQmhQZ01COHQ1dUxJR0dtSk95ei9zc01VN056YVRxQjRsQT09IiwiYW1yIjpbInJzYSIsIm1mYSJdLCJhcHBpZCI6ImQzNTkwZWQ2LTUyYjMtNDEwMi1hZWZmLWFhZDIyOTJhYjAxYyIsImFwcGlkYWNyIjoiMCIsImRldmljZWlkIjoiMjE5OTVlNWYtNDMzOS00M2QyLTg0YTItMjY5Yzg5ZDY3MmVjIiwiZmFtaWx5X25hbWUiOiJBa2JhciIsImdpdmVuX25hbWUiOiJNZCIsImlwYWRkciI6IjI3LjYwLjE0MC45NyIsIm5hbWUiOiJNZCBBa2JhciIsIm9pZCI6ImI4ZGExOTBhLTEwMDYtNDdiYy1iOTZjLWYwZjE0ZDE3ZWQ1NyIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMTQ2NzczMDg1LTkwMzM2MzI4NS03MTkzNDQ3MDctMjI0MDY4NSIsInB1aWQiOiIxMDAzM0ZGRkEyMEJEMkQwIiwicmgiOiIwLkFSb0F2NGo1Y3ZHR3IwR1JxeTE4MEJIYlI5WU9XZE96VWdKQnJ2LXEwaWtxc0J3YUFFOC4iLCJzY3AiOiJBY3Rpdml0eUZlZWQtSW50ZXJuYWwuUmVhZFdyaXRlIENhbGVuZGFycy5SZWFkV3JpdGUgQ29sbGFiLUludGVybmFsLlJlYWQgQ29udGFjdHMuUmVhZFdyaXRlIENvcmVJdGVtLUludGVybmFsLlJlYWQgRVdTLkFjY2Vzc0FzVXNlci5BbGwgRmlsZXMuUmVhZCBGaWxlcy5SZWFkV3JpdGUuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgTWFpbC5SZWFkLkFsbCBNYWlsLlJlYWRXcml0ZSBNYWlsLlNlbmQgTWFwaUh0dHAuQWNjZXNzQXNVc2VyLkFsbCBOb3Rlcy5SZWFkV3JpdGUgTm90ZXMtSW50ZXJuYWwuUmVhZFdyaXRlIE9mZmljZUZlZWQtSW50ZXJuYWwuUmVhZFdyaXRlIE9mZmljZUludGVsbGlnZW5jZS1JbnRlcm5hbC5SZWFkV3JpdGUgUGVvcGxlUHJlZGljdGlvbnMtSW50ZXJuYWwuUmVhZCBQcml2aWxlZ2UuRUxUIFJvYW1pbmdVc2VyU2V0dGluZ3MuUmVhZFdyaXRlIFNpZ25hbHMuUmVhZCBTaWduYWxzLlJlYWRXcml0ZSBTdWJzdHJhdGVTZWFyY2gtSW50ZXJuYWwuUmVhZFdyaXRlIFRhZ3MuUmVhZFdyaXRlIFRhc2tzLlJlYWRXcml0ZSBUb2RvLUludGVybmFsLlJlYWRXcml0ZSB1c2VyX2ltcGVyc29uYXRpb24gVXNlci1JbnRlcm5hbC5SZWFkIiwic3ViIjoibmlQSVk2NXU5eHZNZ3p3OTBJX3d1Y3BHRDk0QU95cGJ0TVVvYkkxcDhvZyIsInRpZCI6IjcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0NyIsInVuaXF1ZV9uYW1lIjoibWRha2JhckBtaWNyb3NvZnQuY29tIiwidXBuIjoibWRha2JhckBtaWNyb3NvZnQuY29tIiwidXRpIjoiLUkydzBmakhCMGlQZlhSNkxtaGxBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il19.h7Guf2nuZC7vYLxJwQKw9unYS0dHPi6b8OX3PzoooJ_KkliXYuHZ2IYel3abEg7kIvBT82O59X75ciISdZRUgBbwmgP6ldp_bAuX7sN3DeWXVlcw_2WuxPj9M2eX3F1pNcWLvFjP7MxeqG2nPrtlKEoIkiIxo5wu9AEL8GY1wMWbdumGLTIhkk13GojrOESDOZV4yFJjqtIzUEIC8-hhm40nmWKC55be4yWQplwk5Ri93DfxOg_ze2ne1aB6NrHgDuwhjfLbapTDAaTfMD9AwHD9BELs_Y-6GURCZfWr4wi8bgaf2oNS3cWletYpiZdBe7YM4RoTRz0GeWx2v8xtXA";
            var preferHeader = "exchange.behavior='SubstrateDefaultFolders, OpenComplexTypeExtensions, ApplicationDataIP'";
            
            var url = "https://substrate.office.com/api/beta/groups/" + groupId + "@" + tenantId + 
                string.Format("/DefaultFolders/TeamsMessagesData/Messages?$filter=ClientThreadId eq '{0}' and ClientConversationId eq '{1}'", threadId, conversationId);
            
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {token}");
            client.DefaultRequestHeaders.Add("Prefer", preferHeader);

            var result = await client.GetAsync(url);

            var customerJsonString = await result.Content.ReadAsStringAsync();

            var deserialized = JsonConvert.DeserializeObject<TeamsMessagesData>(custome‌​rJsonString);

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
