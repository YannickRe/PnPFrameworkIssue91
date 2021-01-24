using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Security;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Providers.Xml;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;

namespace FunctionApp
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log, ExecutionContext context)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            try
            {
                string siteUrl = "";
                string userName = "";
                string password = "";

                var securePassword = new SecureString();
                foreach (char c in password)
                    securePassword.AppendChar(c);
                securePassword.MakeReadOnly();

                var authManager = new AuthenticationManager(userName, securePassword);
                using (ClientContext clientContext = authManager.GetContext(siteUrl))
                {
                    var web = clientContext.Web;

                    var templateFileName = "SitePagesTemplate.pnp";
                    FileConnectorBase fileConnector = new FileSystemConnector(context.FunctionAppDirectory, "");
                    var openXmlConnector = new OpenXMLConnector(templateFileName, fileConnector);
                    var provider = new XMLOpenXMLTemplateProvider(openXmlConnector);
                    templateFileName = templateFileName.Substring(0, templateFileName.LastIndexOf(".", StringComparison.Ordinal)) + ".xml";
                    ProvisioningTemplate provisioningTemplate = provider.GetTemplate(templateFileName);
                    provisioningTemplate.Connector = provider.Connector;

                    var applyingInformation = new ProvisioningTemplateApplyingInformation()
                    {
                        ProgressDelegate = (message, progress, total) =>
                        {
                            log.LogInformation(string.Format("{0:00}/{1:00} - {2}", progress, total, message));
                        },
                        MessagesDelegate = (message, messageType) =>
                        {
                            log.LogInformation(string.Format("{0} - {1}", messageType, message));
                        },
                        IgnoreDuplicateDataRowErrors = true
                    };

                    web.ApplyProvisioningTemplate(provisioningTemplate, applyingInformation);

                    return new OkObjectResult("Done");
                }
            }
            catch (ServerException e)
            {
                log.LogError($"Message: {e.Message}");
                log.LogError($"ServerErrorCode: {e.ServerErrorCode}");
                log.LogError($"ServerErrorDetails: {e.ServerErrorDetails}");
                log.LogError($"ServerErrorTraceCorrelationId: {e.ServerErrorTraceCorrelationId}");
                log.LogError($"ServerErrorTypeName: {e.ServerErrorTypeName}");
                log.LogError($"ServerErrorValue: {e.ServerErrorValue}");
                log.LogError($"ServerStackTrace: {e.ServerStackTrace}");
                log.LogError($"Source: {e.Source}");
                log.LogError($"StackTrace: {e.StackTrace}");

                throw;
            }
            catch (Exception e)
            {
                log.LogError($"Error while processing dossier: {e.Message}\n\n{e.StackTrace}");
                throw;
            }
        }
    }
}
