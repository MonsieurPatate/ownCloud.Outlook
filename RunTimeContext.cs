using System;
using System.IO;
using AdysTech.CredentialManager;
using Microsoft.Office.Interop.Outlook;
using ownCloud.Outlook.Enums;
using RestSharp;
using RestSharp.Authenticators;
using WebDav;

namespace ownCloud.Outlook
{
    /// <summary>
    /// A singleton that contains user info
    /// </summary>
    public class RuntimeContext
    {
        private static RuntimeContext _instance;
        private static IWebDavClient _webDavClient;
        
        // TODO after changing config or credentials recreate refresh client
        private static IWebDavClient WebDavClient
        {
            get
            {
                if (_webDavClient != null)
                {
                    return _webDavClient;
                }

                var config = ConfigManager.Read();
                var credential = CredentialManager.GetCredentials(Constants.AddInName);
                var parameters = new WebDavClientParams
                {
                    BaseAddress = new Uri(config.Server),
                    Credentials = credential,
                };
                _webDavClient = new WebDavClient(parameters);
                return _webDavClient;
            }
        }

        public static RuntimeContext Instance => _instance ?? (_instance = new RuntimeContext());

        public string UploadAttachment(Attachment attachment)
        {
            var credential = CredentialManager.GetCredentials(Constants.AddInName);
            var config = ConfigManager.Read();

            WebDavClient.PutFile(string.Concat($"remote.php/dav/files/{credential.UserName}/", attachment.FileName), File.OpenRead(attachment.GetTemporaryFilePath())).Wait();

            var client = new RestClient($"{config.Server}/ocs/v1.php/apps/files_sharing/api/v1")
            {
                Authenticator = new HttpBasicAuthenticator(credential.UserName, credential.Password)
            };

            var request = new RestRequest("shares")
                .AddParameter("path", "/" + attachment.FileName)
                .AddParameter("shareType", (int) ShareType.Public)
                .AddParameter("permissions", (int) PermissionType.Read)
                .AddParameter("password", 123)
                .AddParameter("name", attachment.FileName);
            var response = client.Post<SharedItem>(request);

            return response.IsSuccessful ? response.Data.Url : $"Could't upload file: {response.ErrorMessage}";
        }
    }
}