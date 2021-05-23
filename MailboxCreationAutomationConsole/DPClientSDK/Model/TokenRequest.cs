using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DPClientSDK.Model
{
	public class TokenRequest
	{
		[JsonProperty(PropertyName = "username")]
		public string Username { get; set; }

		[JsonProperty(PropertyName = "password")]
		public string Password { get; set; }

		[JsonProperty(PropertyName = "refresh_token")]
		public string RefreshToken { get; set; }

		[JsonProperty(PropertyName = "client_id")]
		public string ClientId { get; set; }

		[JsonProperty(PropertyName = "grant_type")]
		public string GrantType { get; set; }
	}
}
