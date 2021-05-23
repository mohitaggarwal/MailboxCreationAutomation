using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DPClientSDK.Model
{
	public class TokenInformation
	{
		[JsonProperty(PropertyName = "access_token")]
		public string AccessToken { get; set; }

		[JsonProperty(PropertyName = "refresh_token")]
		public string RefreshToken { get; set; }

		[JsonProperty(PropertyName = "refresh_expires_in")]
		public int RefreshExpiresIn { get; set; }

		[JsonProperty(PropertyName = "expires_in")]
		public int ExpiresIn { get; set; }

		[JsonIgnore]
		public DateTime TokenExpiresIn { get; set; }

		[JsonProperty(PropertyName = "token_type")]
		public string TokenType { get; set; }

		[JsonProperty(PropertyName = "session_state")]
		public string SessionState { get; set; }

		[JsonProperty(PropertyName = "scope")]
		public string Scope { get; set; }
	}
}
