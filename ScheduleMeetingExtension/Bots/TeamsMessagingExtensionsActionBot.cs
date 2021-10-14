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
using Attachment = Microsoft.Bot.Schema.Attachment;
using File = System.IO.File;
using Microsoft.Graph;
using System.Linq;
using System.Text.RegularExpressions;

namespace Microsoft.BotBuilderSamples.Bots
{
    public class TeamsMessagingExtensionsActionBot : TeamsActivityHandler
    {
        public readonly string baseUrl;
        private string graphToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IkRFX2JUV3BrbFdGcFVKa3hSYzRJUzBMek1UZXJyQ1hPYVhyOUR0dFJxemciLCJhbGciOiJSUzI1NiIsIng1dCI6Imwzc1EtNTBjQ0g0eEJWWkxIVEd3blNSNzY4MCIsImtpZCI6Imwzc1EtNTBjQ0g0eEJWWkxIVEd3blNSNzY4MCJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDcvIiwiaWF0IjoxNjM0MTgzNjA2LCJuYmYiOjE2MzQxODM2MDYsImV4cCI6MTYzNDE4NzUwNiwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsiYzI1Il0sImFpbyI6IkFVUUF1LzhUQUFBQWttRS9qZktabTF0dDgyTUtoeEdoYjlSdG1wT1ludVpBc2FoMFptY3NnV3YrWVkvMGNPTU5DQ3JSMloxQmR0VzBPWEZXZk5WRjk5OU9YcUhHVGU4bVBBPT0iLCJhbXIiOlsicnNhIiwibWZhIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik1pY3Jvc29mdCBPZmZpY2UiLCJhcHBpZCI6ImQzNTkwZWQ2LTUyYjMtNDEwMi1hZWZmLWFhZDIyOTJhYjAxYyIsImFwcGlkYWNyIjoiMCIsImNvbnRyb2xzIjpbImFwcF9yZXMiXSwiY29udHJvbHNfYXVkcyI6WyIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAiXSwiZGV2aWNlaWQiOiIyMTk5NWU1Zi00MzM5LTQzZDItODRhMi0yNjljODlkNjcyZWMiLCJmYW1pbHlfbmFtZSI6IkFrYmFyIiwiZ2l2ZW5fbmFtZSI6Ik1kIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMTE3Ljk2LjE2MC4xNzUiLCJuYW1lIjoiTWQgQWtiYXIiLCJvaWQiOiJiOGRhMTkwYS0xMDA2LTQ3YmMtYjk2Yy1mMGYxNGQxN2VkNTciLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjE0Njc3MzA4NS05MDMzNjMyODUtNzE5MzQ0NzA3LTIyNDA2ODUiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzNGRkZBMjBCRDJEMCIsInJoIjoiMC5BUm9BdjRqNWN2R0dyMEdScXkxODBCSGJSOVlPV2RPelVnSkJydi1xMGlrcXNCd2FBRTguIiwic2NwIjoiQXVkaXRMb2cuUmVhZC5BbGwgQ2FsZW5kYXIuUmVhZFdyaXRlIENhbGVuZGFycy5SZWFkLlNoYXJlZCBDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZSBEYXRhTG9zc1ByZXZlbnRpb25Qb2xpY3kuRXZhbHVhdGUgRGV2aWNlTWFuYWdlbWVudENvbmZpZ3VyYXRpb24uUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudENvbmZpZ3VyYXRpb24uUmVhZFdyaXRlLkFsbCBEaXJlY3RvcnkuQWNjZXNzQXNVc2VyLkFsbCBEaXJlY3RvcnkuUmVhZC5BbGwgRmlsZXMuUmVhZCBGaWxlcy5SZWFkLkFsbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEdyb3VwLlJlYWQuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgTWFpbC5SZWFkV3JpdGUgTm90ZXMuQ3JlYXRlIFBlb3BsZS5SZWFkIFBlb3BsZS5SZWFkLkFsbCBTZW5zaXRpdmVJbmZvVHlwZS5EZXRlY3QgU2Vuc2l0aXZlSW5mb1R5cGUuUmVhZC5BbGwgU2Vuc2l0aXZpdHlMYWJlbC5FdmFsdWF0ZSBVc2VyLlJlYWQuQWxsIFVzZXIuUmVhZEJhc2ljLkFsbCBVc2VyLlJlYWRXcml0ZSBVc2Vycy5SZWFkIiwic2lnbmluX3N0YXRlIjpbImR2Y19tbmdkIiwiZHZjX2NtcCIsImR2Y19kbWpkIiwia21zaSJdLCJzdWIiOiJRRzR0SmRka1RYdnhDQmZKbnBDR0U4X1BtOTIyNWJpbFVNLVgzTHZtSVpJIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IldXIiwidGlkIjoiNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3IiwidW5pcXVlX25hbWUiOiJtZGFrYmFyQG1pY3Jvc29mdC5jb20iLCJ1cG4iOiJtZGFrYmFyQG1pY3Jvc29mdC5jb20iLCJ1dGkiOiJKQ0lOZUxQbllFS2RNZW0wbXg4QUFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX3RjZHQiOjEyODkyNDE1NDd9.jC1OvZpByhlp9zQuJKikOG3wGPH-nQI6N0zq5cwbJk2Z_PcLUqIVQnKJlJmVTFg8EBJ8exld6EANL2ZloR_--aDDTFc_XfbjoHC12Oh-Qs9L0bf9xE9ifBZGLgDa5c8GOQDZ1oDPi6_P4KZPbqE3GQku2EPyRKVDqh5Iu9R5xtUnSJZZdUCMm-1W7szkKVcWR-hXTe0utJfdp07TiKqy947BLqYD_Tofz0UsZoyY1REOngh0y-lA5g8Gkk6vYcuzvRzXtRljMnpKPrdIQX4vFgygL7YVXY7dGWC6W4Y2QvICGIqmBZxZ7vQftPwqzDnfk1r__qWv99PfjwpJvi3OzQ";
        private string substrateToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IlBSV0pJdFRDSVZKUW91Qlhwd0VoN3dZNy04Wk5DazIzNWNTNTR3SUN6bFUiLCJhbGciOiJSUzI1NiIsIng1dCI6Imwzc1EtNTBjQ0g0eEJWWkxIVEd3blNSNzY4MCIsImtpZCI6Imwzc1EtNTBjQ0g0eEJWWkxIVEd3blNSNzY4MCJ9.eyJhdWQiOiJodHRwczovL3N1YnN0cmF0ZS5vZmZpY2UuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3LyIsImlhdCI6MTYzNDE3ODE2NywibmJmIjoxNjM0MTc4MTY3LCJleHAiOjE2MzQxODIwNjcsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJFMllBZ3ZQUE1rck9YYklQa3U0L2Q4SGt5K0VuVWVFVGw2d00vZWEzYzJxaTQrTWttKzZLVXdiSE9waStPRXd0Q0QzODlaZnZRYmsyd3gvZitxMllNK1hONjNJcldnRT0iLCJhbXIiOlsicnNhIiwibWZhIl0sImFwcGlkIjoiZDM1OTBlZDYtNTJiMy00MTAyLWFlZmYtYWFkMjI5MmFiMDFjIiwiYXBwaWRhY3IiOiIwIiwiZGV2aWNlaWQiOiIyMTk5NWU1Zi00MzM5LTQzZDItODRhMi0yNjljODlkNjcyZWMiLCJmYW1pbHlfbmFtZSI6IkFrYmFyIiwiZ2l2ZW5fbmFtZSI6Ik1kIiwiaXBhZGRyIjoiMTE3Ljk2LjE2MC4xNzUiLCJuYW1lIjoiTWQgQWtiYXIiLCJvaWQiOiJiOGRhMTkwYS0xMDA2LTQ3YmMtYjk2Yy1mMGYxNGQxN2VkNTciLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjE0Njc3MzA4NS05MDMzNjMyODUtNzE5MzQ0NzA3LTIyNDA2ODUiLCJwdWlkIjoiMTAwMzNGRkZBMjBCRDJEMCIsInJoIjoiMC5BUm9BdjRqNWN2R0dyMEdScXkxODBCSGJSOVlPV2RPelVnSkJydi1xMGlrcXNCd2FBRTguIiwic2NwIjoiQWN0aXZpdHlGZWVkLUludGVybmFsLlJlYWRXcml0ZSBDYWxlbmRhcnMuUmVhZFdyaXRlIENvbGxhYi1JbnRlcm5hbC5SZWFkIENvbnRhY3RzLlJlYWRXcml0ZSBDb3JlSXRlbS1JbnRlcm5hbC5SZWFkIEVXUy5BY2Nlc3NBc1VzZXIuQWxsIEZpbGVzLlJlYWQgRmlsZXMuUmVhZFdyaXRlLkFsbCBHcm91cC5SZWFkV3JpdGUuQWxsIE1haWwuUmVhZC5BbGwgTWFpbC5SZWFkV3JpdGUgTWFpbC5TZW5kIE1hcGlIdHRwLkFjY2Vzc0FzVXNlci5BbGwgTm90ZXMuUmVhZFdyaXRlIE5vdGVzLUludGVybmFsLlJlYWRXcml0ZSBPZmZpY2VGZWVkLUludGVybmFsLlJlYWRXcml0ZSBPZmZpY2VJbnRlbGxpZ2VuY2UtSW50ZXJuYWwuUmVhZFdyaXRlIFBlb3BsZVByZWRpY3Rpb25zLUludGVybmFsLlJlYWQgUHJpdmlsZWdlLkVMVCBSb2FtaW5nVXNlclNldHRpbmdzLlJlYWRXcml0ZSBTaWduYWxzLlJlYWQgU2lnbmFscy5SZWFkV3JpdGUgU3Vic3RyYXRlU2VhcmNoLUludGVybmFsLlJlYWRXcml0ZSBUYWdzLlJlYWRXcml0ZSBUYXNrcy5SZWFkV3JpdGUgVG9kby1JbnRlcm5hbC5SZWFkV3JpdGUgdXNlcl9pbXBlcnNvbmF0aW9uIFVzZXItSW50ZXJuYWwuUmVhZCIsInN1YiI6Im5pUElZNjV1OXh2TWd6dzkwSV93dWNwR0Q5NEFPeXBidE1Vb2JJMXA4b2ciLCJ0aWQiOiI3MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDciLCJ1bmlxdWVfbmFtZSI6Im1kYWtiYXJAbWljcm9zb2Z0LmNvbSIsInVwbiI6Im1kYWtiYXJAbWljcm9zb2Z0LmNvbSIsInV0aSI6Ik9kdTcycFFLbTA2cThUUlhLR2g5QUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdfQ.JD08foJU2Z1KEujNTv1_ft24oyTx7Z2RYANTLLCON_wQPH0s_-w04ylq1BXoDhQo0GeqMomkNsZD9f6s6_PFFRwOEz_1BxPWhTN6-keHzNo1xYwUQmst5ktKWIvKd1un9iQRjjOv5ZLb_YUltvwkU2JqPQ4B-Qt6s9bgbbVRPI1SOag_gkNsqVxFpzKx8RiDzbddGh9EYfsHR57eNDAlN6uuzKWOCEa4lK4n8nlarCJ9KHKPSPxptJpjNMAbh_XXHvysK5ptISEuoJiH4RfpxO_DtU5nNQutHHDAJAtNd58LG8V2IJYlGBJKr0n_AccRoZ9PgUb0K12Pf1VMmj4MsA";

        public TeamsMessagingExtensionsActionBot(IConfiguration configuration) : base()
        {
            this.baseUrl = configuration["BaseUrl"];
        }

        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            switch (action.CommandId)
            {
                case "scheduleMeeting":
                    return await ScheduleMeetingResponse(turnContext, action);
            }
            return new MessagingExtensionActionResponse();
        }

        private async Task<MessagingExtensionActionResponse> ScheduleMeetingResponse(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        {
            // Fetch data from adaptive card
            var a = action.Data.ToString();
            CustomFormResponse data = JsonConvert.DeserializeObject<CustomFormResponse>(a);
            var title = data.Title;
            var attendeeAddresses = data.AttendeeAddresses.Split(';').AsEnumerable();
            var duration = 30; // default
            var startTime = data.StartTime;

            // Call create meeting api
            var response = await createMeetingAPI(title, attendeeAddresses, startTime, duration);

            // On above api call success
            // Return adaptive card to be inserted into message conversation
            var attachments = new List<MessagingExtensionAttachment>();
            attachments.Add(new MessagingExtensionAttachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>()
                    {
                      new AdaptiveTextBlock() { Text = $"Meeting created with below details - ", Color = AdaptiveTextColor.Accent },
                      new AdaptiveTextBlock() { Text = $"**Title** : {title}" },
                      new AdaptiveTextBlock() { Text = $"**Organizer** : {data.Organizer}" },
                      new AdaptiveTextBlock() { Text = $"**Attendees** : {data.AttendeeNames}" },
                      new AdaptiveTextBlock() { Text = $"**Time** : {startTime.ToString("dddd, dd MMMM yyyy hh:mm tt")}" },
                      new AdaptiveTextBlock() { Text = $"**Duration** : {data.Duration}" },
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

        private async Task<Event> createMeetingAPI(string title, IEnumerable<string> attendeeAddresses, DateTime startTime, int duration)
        {
            var scopes = new[] { "User.Read" };

            var authProvider = new DelegateAuthenticationProvider(async (request) => {
                // Use Microsoft.Identity.Client to retrieve token
                // var result = await pca.AcquireTokenByIntegratedWindowsAuth(scopes).ExecuteAsync();

                var token = graphToken;
                request.Headers.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            });

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            var @event = new Event
            {
                Subject = title,
                //Body = new ItemBody
                //{
                //    ContentType = BodyType.Html,
                //    Content = "Does noon work for you?"
                //},
                Start = new DateTimeTimeZone
                {
                    DateTime = startTime.ToString(),
                    TimeZone = "India Standard Time"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = startTime.AddMinutes(duration).ToString(),
                    TimeZone = "India Standard Time"
                },
                //Location = new Location
                //{
                //    DisplayName = "online meet"
                //},
                Attendees = attendeeAddresses.Select(a => {
                    return new Attendee
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = a,
                        },
                        Type = AttendeeType.Required
                    };
                })
            };
            var response = await graphClient.Me.Events
                .Request()
                .Header("Prefer", "outlook.timezone=\"India Standard Time\"")
                .AddAsync(@event);

            var x = response.Subject;
            return response;
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
            // Populate from request
            Regex regex = new Regex(@".*;messageid=(.*)");
            var match = regex.Match(turnContext.Activity.Conversation.Id);
            var parentMessageid = match.Groups[1].Value;// "1634033527688";
            var organizer = new EmailAddress()
            {
                Name = turnContext.Activity.From.Name,
                Address = "" // Setting it below, as we only have name in turn context.
            };
            var groupId = "b77dde63-1f70-4668-95d5-0c34e038514b";
            var channelData = (JObject)turnContext.Activity.ChannelData;
            var threadId = channelData["channel"]["id"].Value<string>();// "19:dbad371212c0450f8c63da35ef7a0484@thread.tacv2";
            var tenantId = turnContext.Activity.Conversation.TenantId;// "72f988bf-86f1-41af-91ab-2d7cd011db47";
            var conversationId = turnContext.Activity.Conversation.Id;// "19:dbad371212c0450f8c63da35ef7a0484@thread.tacv2;messageid=1634033527688";

            // Get all conversation messages
            var allMessages = await GetAllMessages(groupId, threadId, conversationId, tenantId);

            // Fetch coversation title, attendees
            var title = allMessages.Value.Where(v => v.InternetMessageId == parentMessageid).FirstOrDefault().Subject;
            var a = allMessages.Value.GroupBy(v => v.Sender.EmailAddress.Address);
            var b = a.Select(v => v.FirstOrDefault());
            var allAttendees = b.Select(v => {
                return new AttendeeBase
                {
                    Type = AttendeeType.Required,
                    EmailAddress = new EmailAddress
                    {
                        Address = v.Sender.EmailAddress.Address,
                        Name = v.Sender.EmailAddress.Name
                    }
                };
            });
            var attendees = allAttendees.Where(a => {
                if (a.EmailAddress.Name != organizer.Name)
                {
                    return true;
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(organizer.Address))
                    {
                        organizer.Address = a.EmailAddress.Address;
                    }
                    return false;
                }
            });

            // Call FMT api to get meeting start time
            var fmtResponse = await FMT(allAttendees);
            var startTime = DateTime.Parse(fmtResponse.MeetingTimeSuggestions.ToList()[0].MeetingTimeSlot.Start.DateTime);

            // Return meeting details adaptive card
            var response = new MessagingExtensionActionResponse()
            {
                Task = new TaskModuleContinueResponse()
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Height = 600,
                        Width = 450,
                        Title = "Meeting details",
                        Card = GetAdaptiveCardAttachmentForMeeting(title, attendees, organizer, startTime),
                    },
                },
            };
            return response;
        }

        private async Task<TeamsMessagesData> GetAllMessages(string groupId, string threadId, string conversationId, string tenantId)
        {
            var token = substrateToken;

            // Get conversation Messages
            var preferHeader = "exchange.behavior='SubstrateDefaultFolders, OpenComplexTypeExtensions, ApplicationDataIP'";

            var url = "https://substrate.office.com/api/beta/groups/" + groupId + "@" + tenantId +
                string.Format("/DefaultFolders/TeamsMessagesData/Messages?$filter=ClientThreadId eq '{0}' and ClientConversationId eq '{1}'", threadId, conversationId);

            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {token}");
            client.DefaultRequestHeaders.Add("Prefer", preferHeader);

            var result = await client.GetAsync(url);

            var customerJsonString = await result.Content.ReadAsStringAsync();

            return JsonConvert.DeserializeObject<TeamsMessagesData>(custome‌​rJsonString);
        }

        private async Task<MeetingTimeSuggestionsResult> FMT(IEnumerable<AttendeeBase> allAttendees)
        {
            var authProvider = new DelegateAuthenticationProvider(async (request) => {
                var token = graphToken;
                request.Headers.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            });
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var locationConstraint = new LocationConstraint
            {
                IsRequired = false,
                SuggestLocation = false,
                Locations = new List<LocationConstraintItem>()
                {
                    new LocationConstraintItem
                    {
                        ResolveAvailability = false,
                        DisplayName = "Conf room Hood"
                    }
                }
            };

            var timeConstraint = new TimeConstraint
            {
                ActivityDomain = ActivityDomain.Work,
                TimeSlots = new List<TimeSlot>()
                {
                    new TimeSlot
                    {
                        Start = new DateTimeTimeZone
                        {
                            DateTime = DateTime.Now.ToString(),
                            TimeZone = "India Standard Time"
                        },
                        End = new DateTimeTimeZone
                        {
                            DateTime = DateTime.Now.AddDays(7).ToString(),
                            TimeZone = "India Standard Time"
                        }
                    }
                }
            };

            var isOrganizerOptional = false;
            var meetingDuration = new Duration("PT30M");
            var returnSuggestionReasons = true;
            var minimumAttendeePercentage = (double)100;

            return await graphClient.Me
                .FindMeetingTimes(allAttendees, locationConstraint, timeConstraint, meetingDuration, null, isOrganizerOptional, returnSuggestionReasons, minimumAttendeePercentage)
                .Request()
                .Header("Prefer", "outlook.timezone=\"India Standard Time\"")
                .PostAsync();
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

        private static Attachment GetAdaptiveCardAttachmentForMeeting(string title, IEnumerable<AttendeeBase> attendees, EmailAddress organizer, DateTime startTime, int duration = 30)
        {
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>()
                    {
                      new AdaptiveTextInput() { Id = "title", Value = title, Label = "Title" },
                      new AdaptiveTextInput() { Id = "organizer", Value = organizer.Name, Label = "Organizer" },
                      new AdaptiveTextInput() { Id = "attendeeNames", Value = string.Join("; ", attendees.Select(a => a.EmailAddress.Name)), Label = "Attendees" },
                      new AdaptiveTextInput() { Id = "attendeeAddresses", Value = string.Join("; ", attendees.Select(a => a.EmailAddress.Address)), IsVisible = false},
                      new AdaptiveTextInput() { Id = "startTime", Value = startTime.ToString("dddd, dd MMMM yyyy hh:mm tt"), Label = "Start time" },
                      new AdaptiveTextInput() { Id = "duration", Value = $"{duration.ToString()} minute", Label = "Duration" },
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
