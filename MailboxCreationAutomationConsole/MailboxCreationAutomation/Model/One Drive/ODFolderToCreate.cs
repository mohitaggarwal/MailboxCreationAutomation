using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class ODFolderToCreate
	{
		public string Prefix { get; set; }
		public int Count { get; set; }
		public int Levels { get; set; }
		public List<ODFileToCreate> FilesToCreate { get; set; }
	}
}
