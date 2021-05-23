using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class UserToCreate
	{
		public string DisplayName { get; set; }
		public string MailNickName { get; set; }
		public string UserPrinicpleName { get; set; }
		public string Password { get; set; }
		public string UsageLocation { get; set; }
		public string LicenseSkuId { get; set; }
	}

}
