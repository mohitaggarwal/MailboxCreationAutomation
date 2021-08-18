using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class OneDriveToCreate
	{
		public List<ODFileToCreate> RootFilesToCreate { get; set; }
		public List<ODFolderToCreate> FoldersToCreate { get; set; }
	}
}
