using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class MailsToCreate
	{
		public List<string> To { get; set; }
		public string From { get; set; }
		public string Subject { get; set; }
		public string BodyPath { get; set; }
		public int Count { get; set; }
		public List<AttachmentsToCreate> AttachmentsToCreateList { get; set; }
	}
}
