using Azure.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace appsvc_fnc_scw_unclassified
{
 
        public class ROPCConfidentialTokenCredential : Azure.Core.TokenCredential
        {
            // Implementation of the Azure.Core.TokenCredential class
            string _username = "";
            string _password = "";
            string _tenantId = "";
            string _clientId = "";
            string _clientSecret = "";

            string _tokenEndpoint = "";

            public ROPCConfidentialTokenCredential(string username, string password, string tenantId, string clientId, string clientSecret)
            {
                // Public Constructor
                _username = username;
                _password = password;
                _tenantId = tenantId;
                _clientId = clientId;
                _clientSecret = clientSecret;

                _tokenEndpoint = "https://login.microsoftonline.com/" + _tenantId + "/oauth2/v2.0/token";
            }

            public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
            {
                HttpClient httpClient = new HttpClient();

                // Create the request body
                var Parameters = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_secret", _clientSecret),
                new KeyValuePair<string, string>("scope", string.Join(" ", requestContext.Scopes)),
                new KeyValuePair<string, string>("username", _username),
                new KeyValuePair<string, string>("password", _password),
                new KeyValuePair<string, string>("grant_type", "password")
            };

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenEndpoint)
                {
                    Content = new FormUrlEncodedContent(Parameters)
                };
                var response = httpClient.SendAsync(request).Result.Content.ReadAsStringAsync().Result;
                dynamic responseJson = JsonConvert.DeserializeObject(response);
                var expirationDate = DateTimeOffset.UtcNow.AddMinutes(60.0);
                return new AccessToken(responseJson.access_token.ToString(), expirationDate);
            }

            public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
            {
                HttpClient httpClient = new HttpClient();

                // Create the request body
                var Parameters = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_secret", _clientSecret),
                new KeyValuePair<string, string>("scope", string.Join(" ", requestContext.Scopes)),
                new KeyValuePair<string, string>("username", _username),
                new KeyValuePair<string, string>("password", _password),
                new KeyValuePair<string, string>("grant_type", "password")
            };

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenEndpoint)
                {
                    Content = new FormUrlEncodedContent(Parameters)
                };
                var response = httpClient.SendAsync(request).Result.Content.ReadAsStringAsync().Result;
                dynamic responseJson = JsonConvert.DeserializeObject(response);
                var expirationDate = DateTimeOffset.UtcNow.AddMinutes(60.0);
                return new ValueTask<AccessToken>(new AccessToken(responseJson.access_token.ToString(), expirationDate));
            }
       // }
    }
}
