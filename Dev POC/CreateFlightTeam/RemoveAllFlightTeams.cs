// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
using MIPActivateMSGraphAPI_CreateFlightTeam.Graph;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;

namespace MIPActivateMSGraphAPI_CreateFlightTeam
{
    public static class RemoveAllFlightTeams
    {
        private static ILogger logger = null;

        [FunctionName("RemoveAllFlightTeams")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequest req, ILogger log)
        {
            logger = log;

            try
            {
                // Exchange token for Graph token via on-behalf-of flow
                var graphToken = await AuthProvider.GetTokenOnBehalfOfAsync(req.Headers["Authorization"]);
                logger.LogInformation($"Access token: {graphToken}");

                await RemoveAllTeamsAsync(graphToken);

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

        public static async Task RemoveAllTeamsAsync(string accessToken)
        {
            // Initialize Graph client
            var graphClient = new GraphService(accessToken, logger);

            bool more = true;

            do
            {
                var groups = await graphClient.GetAllGroupsAsync("startswith(displayName, 'Flight')");

                more = groups.Value.Count > 0;

                foreach (var group in groups.Value)
                {
                    if (group.DisplayName != "Flight Admin")
                    {
                        logger.LogInformation($"Deleting team {group.DisplayName}");

                        // Archive the team
                        try
                        {
                            await graphClient.ArchiveTeamAsync(group.Id);
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Contains("ItemNotFound"))
                            {
                                logger.LogInformation("No team found");
                            }
                            else { throw ex; }
                        }

                        // Delete the group
                        await graphClient.DeleteGroupAsync(group.Id);
                    }
                }
            }
            while (more);
        }
    }
}
