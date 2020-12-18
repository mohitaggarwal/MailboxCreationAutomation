using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class MailsToCreate
	{
		public string From { get; set; }
		public int Mails { get; set; }
		public int Attachments { get; set; }
		public int AttachmentSizeInMB { get; set; }
	}
}
