using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class Drive
	{
		private GraphServiceWrapper _GraphServiceWrapper;

		public Drive(GraphServiceWrapper graphServiceWrapper)
		{
			_GraphServiceWrapper = graphServiceWrapper;
		}

		public string GetDriveId(string userId)
		{
			var drive = _GraphServiceWrapper.GraphServiceClient
								.Users[userId]
								.Drive
								.Request()
								.GetAsync().Result;
			return drive.Id;
		}
	}
}
