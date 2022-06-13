using Microsoft.Office.Insights.DataFactory.Activities.AutobuggingServiceBus.ServiceBusUtilities.Messages;
using MS.Office.Utilities.Email;

namespace Microsoft.Office.Telemetry.DataFactory.Activities.Alerting.ServiceBusManager
{
    public static class EmailProcessingUtilities
    {
        private static MicrosoftGraphClient GraphClient;
        private static object GraphClientCreationLock = new object();

        public static void CreateGraphClient()
        {
            var graphClientCreationTask = MicrosoftGraphClient.CreateGraphClientUsingDelegatedPermissions
                        (AssemblyConfig.GetValue("EmailTenantID"), AssemblyConfig.GetValue("EmailClientID"),
                        AssemblyConfig.GetValue("EmailAuthenticationUser"),
                        AssemblyConfig.GetValue("EmailAuthenticationPass"));
            GraphClient = graphClientCreationTask.Result;
        }

        public static void SendEmailMessage(EmailMessage emailMessage)
        {
            lock (GraphClientCreationLock)
            {
                if (GraphClient == null)
                {
                    CreateGraphClient();
                }
            }

            GraphClient.SendEmail(emailMessage.Body, emailMessage.Subject, emailMessage.ToRecipients, bodyType: Graph.BodyType.Html, saveToSentItems: true,
                CC: emailMessage.CcRecipients, BCC: emailMessage.BccRecipients).Wait();
        }
    }
}
