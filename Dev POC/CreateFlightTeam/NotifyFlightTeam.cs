// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
using MIPActivateMSGraphAPI_CreateFlightTeam.Graph;
using MIPActivateMSGraphAPI_CreateFlightTeam.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace MIPActivateMSGraphAPI_CreateFlightTeam
{
    public static class NotifyFlightTeam
    {
        private static readonly string notifAppId = Environment.GetEnvironmentVariable("NotificationAppId");
        private static readonly bool sendCrossDeviceNotifications = !string.IsNullOrEmpty(notifAppId);

        private static ILogger logger = null;

        [FunctionName("NotifyFlightTeam")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)]HttpRequest req, ILogger log)
        {
            logger = log;

            try
            {
                // Exchange token for Graph token via on-behalf-of flow
                var graphToken = await AuthProvider.GetTokenOnBehalfOfAsync(req.Headers["Authorization"]);
                logger.LogInformation($"Access token: {graphToken}");

                string requestBody = new StreamReader(req.Body).ReadToEnd();
                var request = JsonConvert.DeserializeObject<NotifyFlightTeamRequest>(requestBody);

                await NotifyTeamAsync(graphToken, request);

                return new OkResult();
            }
            catch (MsalException ex)
            {
                logger.LogInformation($"Could not obtain Graph token: {ex.Message}");
                // Just return 401 if something went wrong
                // during token exchange
                return new UnauthorizedResult();
            }
            catch (Exception ex)
            {
                logger.LogInformation($"Exception occured: {ex.Message}");
                return new BadRequestObjectResult(ex);
            }
        }

        private static async Task NotifyTeamAsync(string accessToken, NotifyFlightTeamRequest request)
        {
            // Initialize Graph client
            var graphClient = new GraphService(accessToken, logger);

            // Find groups with specified SharePoint item ID
            var groupsToNotify = await graphClient.FindGroupsBySharePointItemIdAsync(request.SharePointItemId);

            foreach (var group in groupsToNotify.Value)
            {
                // Post a Teams chat
                await PostTeamChatNotification(graphClient, group.Id, request.NewDepartureGate);

                /* Below code requires extra configuration to work.
                if (sendCrossDeviceNotifications)
                {
                    // Get the group members
                    var members = await graphClient.GetGroupMembersAsync(group.Id);

                    // Send notification to each member
                    await SendNotificationAsync(graphClient, members.Value, group.DisplayName, request.NewDepartureGate);
                }
                */
            }
        }

        private static async Task PostTeamChatNotification(GraphService graphClient, string groupId, string newDepartureGate)
        {
            // Get channels
            var channels = await graphClient.GetTeamChannelsAsync(groupId);

            var notificationMessage = new ChatMessage
            {
                Body = new ItemBody { Content = $"Your flight will now depart from gate {newDepartureGate}" }
            };

            // Post to all channels
            foreach (var channel in channels.Value)
            {
                await graphClient.CreateChatMessageAsync(groupId, channel.Id, notificationMessage);
            }
        }

        private static async Task SendNotificationAsync(GraphService graphClient, List<User> users, string groupName, string newDepartureGate)
        {
            // Ideally loop through all the members here and send each a notification
            // The notification API is currently limited to only send to the logged-in user
            // So to do this, would need to manage tokens for each user.
            // For now, just send to the authenticated user.
            var notification = new Notification
            {
                TargetHostName = notifAppId,
                AppNotificationId = "testDirectToastNotification",
                GroupName = "TestGroup",
                ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(1).ToUniversalTime(),
                Priority = "High",
                DisplayTimeToLive = 30,
                Payload = new NotificationPayload
                {
                    VisualContent = new NotificationVisualContent
                    {
                        Title = $"{groupName} gate change",
                        Body = $"Departure gate has been changed to {newDepartureGate}"
                    }
                },
                TargetPolicy = new NotificationTargetPolicy
                {
                    PlatformTypes = new string[] { "windows", "android", "ios" }
                }
            };

            await graphClient.SendNotification(notification);
        }
    }
}
