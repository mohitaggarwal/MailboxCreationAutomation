using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class OAuthInfo
	{
		public string ClientId { get; set; }
		public string ClientSecret { get; set; }
		public string TenantId { get; set; }
		public string ImpersonateUser { get; set; }
	}

	public class BasicAuthInfo
	{
		public string Username { get; set; }
		public string Password { get; set; }
	}
}
