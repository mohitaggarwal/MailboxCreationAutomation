using MailboxCreationAutomation.Model;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class GraphServiceWrapper
	{
		public GraphServiceClient GraphServiceClient { get; }

		public OAuthInfo OAuthInfo { get; }

		public GraphServiceWrapper(OAuthInfo oAuthInfo)
		{
			OAuthInfo = oAuthInfo;
			IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
																			.Create(oAuthInfo.ClientId)
																			.WithTenantId(oAuthInfo.TenantId)
																			.WithClientSecret(oAuthInfo.ClientSecret)
																			.Build();
			List<string> scopes = new List<string>();
			scopes.Add("https://graph.microsoft.com/.default");
			GraphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) => {

														// Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
														var authResult = await confidentialClientApplication.AcquireTokenForClient(scopes).ExecuteAsync();

														// Add the access token in the Authorization header of the API
														requestMessage.Headers.Authorization =
														new AuthenticationHeaderValue("Bearer", authResult.AccessToken);

																})
														);
		}

	}
}
