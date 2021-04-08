using System;
using System.IO;
using System.Net;
using AdysTech.CredentialManager;
using Microsoft.Office.Interop.Outlook;
using ownCloud.Outlook.Enums;
using RestSharp;
using RestSharp.Authenticators;
using WebDav;

namespace ownCloud.Outlook
{
    /// <summary>
    /// A singleton that manages everything
    /// </summary>
    public class RunTimeContext
    {
        private static RunTimeContext _instance;
        private static IWebDavClient _webDavClient;

        public static RunTimeContext Instance => _instance ?? (_instance = new RunTimeContext());

        public void Init()
        {
            var credential = CredentialManager.GetICredential(Constants.AddInName)?.ToNetworkCredential();
            if (credential != null)
            {
                bool save = true;
                credential = CredentialManager.PromptForCredentials(Constants.AddInName, ref save, "Please, enter your credentials to log in", "Credentials for ownCloud.Outlook AddIn");
            }

            //CredentialManager.SaveCredentials("ownCloud.Outlook AddIn", credential);

            var parameters = new WebDavClientParams
            {
                BaseAddress = new Uri("url"),
                Credentials = new NetworkCredential("login", "pass")
            };
            _webDavClient = new WebDavClient(parameters);
        }

        public string UploadAttachment(Attachment attachment)
        {


            _webDavClient.PutFile(string.Concat("remote.php/dav/files/login/", attachment.FileName), File.OpenRead(attachment.GetTemporaryFilePath())).Wait();

            var client = new RestClient("url/ocs/v1.php/apps/files_sharing/api/v1")
            {
                Authenticator = new HttpBasicAuthenticator("login", "pass")
            };

            var request = new RestRequest("shares")
                .AddParameter("path", "/" + attachment.FileName)
                .AddParameter("shareType", ShareType.Public)
                .AddParameter("permissions", PermissionType.Read)
                .AddParameter("password", 123)
                .AddParameter("name", attachment.FileName);
            var response = client.Post<SharedItem>(request);

            return response.IsSuccessful ? response.Data.Url : $"Could't upload file: {response.ErrorMessage}";
        }
    }
}