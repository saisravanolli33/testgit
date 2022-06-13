using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace MS.Office.Utilities.Email
{
    //TODO: The Custom activity jobs run on machines which don't come with the required version of .net core, so
    // a MethodNotFound exception is thrown. Once that is fixed this copy of the class can go away.
    public class MicrosoftGraphClient
    {
        private HttpProvider httpProvider;
        private GraphServiceClient serviceClient;
        private string sender;

        private MicrosoftGraphClient(IAuthenticationProvider authenticationProvider, string sender)
        {
            serviceClient = new GraphServiceClient(authenticationProvider);
            ErrorHandling.VerifyStringNotNullNorWhiteSpace(sender, "sender");
            this.sender = sender;
        }

        private MicrosoftGraphClient(string accessToken, string sender)
        {
            httpProvider = new HttpProvider();
            serviceClient = new GraphServiceClient
            (
                new DelegateAuthenticationProvider
                (
                    requestMessage =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                        return Task.FromResult(0);
                    }
                ),
                httpProvider
            );
            ErrorHandling.VerifyStringNotNullNorWhiteSpace(sender, "sender");
            this.sender = sender;
        }

        /// <summary>
        /// Use this to create a client if your service principal has Application level Mail.Send permissions i.e.
        /// it can send emails as anyone in the tenant.
        /// Sender should be of the form: alias@microsoft.com
        /// </summary>
        /// <returns></returns>
        public static MicrosoftGraphClient CreateGraphClientUsingApplicationPermissions(string tenantId, string clientId, string clientSecret, string sender)
        {
            var publicClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();
            var authProvider = new ClientCredentialProvider(publicClientApplication);

            return new MicrosoftGraphClient(authProvider, sender);
        }

        /// <summary>
        /// Use this to create a client if your service principal has Mail.Send permissions at a delegated level i.e.
        /// it can only send emails on signed in users, in this case service accounts. Note that per:
        /// https://microsoft.sharepoint.com/sites/mywork/SitePages/Email/Sending-Email-with-Modern-Authentication.aspx
        /// the code needs to be executing on Azure or on corpnet, so logins for service accounts will fail through VPN.
        /// Sender should be a service account and of the form alias@microsoft.com.
        /// </summary>
        /// <returns></returns>

        public static async Task<MicrosoftGraphClient> CreateGraphClientUsingDelegatedPermissions(string tenantId, string clientId, string sender, string senderPassword)
        {
            // UserPasswordCredential is not supported in .NET standard or .NET core, so have to manually create the HTTP post request to get an authorization token.
            // TODO: Test this logic woud when permissions propagate through in Azure.
            using (HttpClient client = new HttpClient())
            {
                var tokenEndpoint = $"https://login.microsoftonline.com/{tenantId}/oauth2/token";

                client.DefaultRequestHeaders.Add("Accept", "application/json");
                string postBody = $"resource=https://graph.microsoft.com&client_id={clientId}&grant_type=password&username={sender}&password={senderPassword}";
                using (HttpResponseMessage response = await client.PostAsync(tokenEndpoint, new StringContent(postBody, Encoding.UTF8, "application/x-www-form-urlencoded")))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        var jsonresult = JObject.Parse(await response.Content.ReadAsStringAsync());
                        var token = (string)jsonresult["access_token"];
                        return new MicrosoftGraphClient(token, sender);
                    }
                    else
                    {
                        throw new UnauthorizedAccessException(string.Format("Unable to obtain an access token from Microsoft Graph API. Ensure that the client ID has delegated and admin consented Mail.Send " +
                            "permissions and that the service account is setup as domain mailbox account. Authentication must also take place inside corpnet or from an Azure VM. Http response: {0}",
                            response.Content.ReadAsStringAsync()));
                    }
                }
            }
        }

        /// <summary>
        /// A simple method to send email using the provided parameters, optionally allowing attachments.
        /// </summary>
        /// <returns></returns>

        public async Task SendEmail
        (
            string content,
            string subject,
            List<string> recipients,
            IMessageAttachmentsCollectionPage attachments = null,
            BodyType bodyType = BodyType.Html,
            Importance? importance = null,
            bool saveToSentItems = false,
            List<string> CC = null,
            List<string> BCC = null
        )
        {
            ErrorHandling.VerifyStringNotNullNorWhiteSpace(content, "content");
            ErrorHandling.VerifyStringNotNullNorWhiteSpace(subject, "subject");
            List<Recipient> emailRecipients = new List<Recipient>();
            List<Recipient> ccRecipients = new List<Recipient>();
            List<Recipient> bccRecipients = new List<Recipient>();

            foreach (var recipient in recipients)
            {
                ErrorHandling.VerifyStringNotNullNorWhiteSpace(recipient, "recipient");
                emailRecipients.Add(new Recipient() { EmailAddress = new EmailAddress() { Address = recipient } });
            }
            var message = new Message()
            {
                Body = new ItemBody { Content = content, ContentType = bodyType },
                Sender = new Recipient() { EmailAddress = new EmailAddress() { Address = sender } },
                ToRecipients = emailRecipients,
                Subject = subject
            };
            if (importance.HasValue)
            {
                message.Importance = importance.Value;
            }
            if (attachments != null && attachments.Count > 0)
            {
                message.Attachments = attachments;
                message.HasAttachments = true;
            }

            if (CC != null && CC.Count > 0)
            {
                foreach (var recipient in CC)
                {
                    ccRecipients.Add(new Recipient() { EmailAddress = new EmailAddress() { Address = recipient } });
                }
                message.CcRecipients = ccRecipients;
            }

            if (BCC != null && BCC.Count > 0)
            {
                foreach (var recipient in BCC)
                {
                    bccRecipients.Add(new Recipient() { EmailAddress = new EmailAddress() { Address = recipient } });
                }
                message.BccRecipients = bccRecipients;
            }
            await serviceClient.Users[sender].SendMail(message, saveToSentItems).Request().PostAsync();
        }

        /// <summary>
        /// Use this if there is a need for a lot of customization for the email message e.g. cc, bcc recipients, read receipts etc. Construct the message
        /// outside of the class and then use the method to send.
        /// </summary>
        /// <returns></returns>

        public async Task SendEmail(Message message, bool saveToSentItems = false)
        {
            await serviceClient.Users[sender].SendMail(message, saveToSentItems).Request().PostAsync();
        }
    }
}
