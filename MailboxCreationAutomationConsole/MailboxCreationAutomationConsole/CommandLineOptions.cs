using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomationConsole
{
	public class CommandLineOptions
	{
		public string AuthenticationType { get; set; }
		public string Username { get; set; }
		public string Password { get; set; }
		public string ClientId { get; set; }
		public string ClientSecret { get; set; }
		public string TenantId { get; set; }
		public string ImpersonateUser { get; set; }
		public string JsonFile { get; set; }
		public bool ShowHelp { get; set; }
		public string  DeleteUser { get; set; }
		public bool CreateUser { get; set; }

	}
}
