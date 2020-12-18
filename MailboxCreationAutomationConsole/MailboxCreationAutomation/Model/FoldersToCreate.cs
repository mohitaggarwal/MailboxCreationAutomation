using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class FoldersToCreate
	{
		public string RootFolder { get; set; }
		public int Folders { get; set; }
		public int Levels { get; set; }

		public MailsToCreate MailsToCreate { get; set; }
	}
}
