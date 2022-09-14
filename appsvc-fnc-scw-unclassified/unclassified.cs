using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Collections.Generic;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;

namespace appsvc_fnc_scw_unclassified
{
    public static class unclassified
    {
        [FunctionName("unclassified")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            IConfiguration config = new ConfigurationBuilder()

           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            var groupid = config["groupid"];
            var labelid = config["labelid"];

            var scopes = new[] { "AuditLog.Read.All", "Directory.ReadWrite.All" };
        
            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential("", "", "", "", "");
            
            var graphClient = new GraphServiceClient(auth, scopes);
            addunclassifiedlabel(graphClient, labelid, groupid, log);

            string responseMessage = "True";

            return new OkObjectResult(responseMessage);
        }

        public static async Task<string> addunclassifiedlabel(GraphServiceClient graphClient, string labelid, string groupid, ILogger log)
        {
            var group = new Group
            {
                AssignedLabels = new List<AssignedLabel>()
                    {
                        new AssignedLabel
                        {
                            LabelId = labelid
                        }
                    },
            };

            var users = await graphClient.Groups[groupid].Request().UpdateAsync(group);
            return "true";
        }
    }
}
