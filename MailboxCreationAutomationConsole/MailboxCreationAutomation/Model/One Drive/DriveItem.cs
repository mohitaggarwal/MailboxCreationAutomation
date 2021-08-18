using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public enum DriveItemType
	{
		File,
		Folder,
		Other
	}

	public class DriveItemInfo
	{
		public string Id { get; set; }
		public string Name { get; set; }
		public long? Size { get; set; }

		public string CheckSum { get; set; }
		public string DownloadUrl { get; set; }
		public DriveItemType DriveItemType { get; set; }
	}
}
