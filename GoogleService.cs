using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

namespace DrinksCaculator
{
    class GoogleService
    {
        public UserCredential Credential { get; set; }

        public string[] Scopes { get; set; }

        public string ApplicationName { get { return "DrinksCaculator"; } }

        public GoogleService(string GoogleCrednetialLoaction)
        {
            using (var stream = new FileStream(GoogleCrednetialLoaction, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                Credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
        }
    }
}
