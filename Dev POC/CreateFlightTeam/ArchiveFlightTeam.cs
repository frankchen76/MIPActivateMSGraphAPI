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
using System.IO;
using System.Threading.Tasks;

namespace MIPActivateMSGraphAPI_CreateFlightTeam
{
    public static class ArchiveFlightTeam
    {
        private static ILogger logger = null;

        [FunctionName("ArchiveFlightTeam")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)]HttpRequest req, ILogger log)
        {
            logger = log;

            try
            {
                // Exchange token for Graph token via on-behalf-of flow
                var graphToken = await AuthProvider.GetTokenOnBehalfOfAsync(req.Headers["Authorization"]);
                logger.LogInformation($"Access token: {graphToken}");

                string requestBody = new StreamReader(req.Body).ReadToEnd();
                var request = JsonConvert.DeserializeObject<ArchiveTeamRequest>(requestBody);

                await ArchiveTeamAsync(graphToken, request);

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

        private static async Task ArchiveTeamAsync(string accessToken, ArchiveTeamRequest request)
        {
            // Initialize Graph client
            var graphClient = new GraphService(accessToken, logger);

            // Find groups with the specified SharePoint item ID
            var groupsToArchive = await graphClient.FindGroupsBySharePointItemIdAsync(request.SharePointItemId);

            foreach (var group in groupsToArchive.Value)
            {
                // Archive each matching team
                await graphClient.ArchiveTeamAsync(group.Id);
            }
        }
    }
}
