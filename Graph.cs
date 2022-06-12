using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using EdiUtilities.ConfigurationChangeGates;
using Microsoft.Graph;
using MS.Office.SharedUtilities.Email;
using MS.Office.SharedUtilities.Email.Models;
using OExp_ToolsSDK.Common.Extensions;

namespace EdiUtilities
{
    public static class EdiEmail
    {
	    private static MicrosoftGraphClient GetClient()
        {
            return MicrosoftGraphClient.CreateGraphClientUsingDelegatedPermissions(
	            tenantId: "72f988bf-86f1-41af-91ab-2d7cd011db47",
	            clientId: "8cd517f3-1be3-4786-8fde-8adb6baaccce" /*Metrical Alerts*/,
	            sender: "edibot@microsoft.com",
	            senderPassword: AzureKeyVault.GetSecretValue("EdiBotPassword")).Result;
        }

        public static bool Send(string content, string subject, List<string> recipients, BodyType bodyType = BodyType.Html, IMessageAttachmentsCollectionPage attachments = null, bool saveToSentItems = false, List<string> CC = null, Importance? importance = null, Dictionary<string, EmbeddedImage> embeddedImages = null)
        { 
            if (EdiEnvironment.IsProduction)
            {
                try 
                { 
                    MicrosoftGraphClient client = GetClient();
                    List<string> resolvedToRecipients = recipients.GetEmailIds().ToList();
                    List<string> resolvedCcRecipients = CC.GetEmailIds().ToList();
                    bool result = HttpClientExtensions.RetryExecutorAsync(() =>
                    {
                        client.SendEmail(content, subject, resolvedToRecipients, bodyType, attachments, saveToSentItems, resolvedCcRecipients, importance, embeddedImages);
                        return Task.FromResult(true);
                    }).GetAwaiter().GetResult();

                }
                catch (Exception ex) 
                {
                    EdiLogger.Error($"Unable to send email titled '{subject}'. Exception: '{ex}'");
                    return false;
                }
            }

            return true;
        }
    }
}
