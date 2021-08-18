using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class OneDrive
	{
		private GraphServiceWrapper _GraphServiceWrapper;
		private string _userId;
		private OneDriveToCreate _OneDriveToCreate;

		public OneDrive(GraphServiceWrapper graphServiceWrapper, OneDriveToCreate oneDriveToCreate,string userId)
		{
			_GraphServiceWrapper = graphServiceWrapper;
			_OneDriveToCreate = oneDriveToCreate;
			_userId = userId;
		}

		public void CreateOneDrive()
		{
			if (_OneDriveToCreate != null && _GraphServiceWrapper != null)
			{
				Logger.FileLogger.Info($"OneDrive '{_userId}' creation started ...");
				Console.WriteLine($"OneDrive '{_userId}' creation started ...");

				ODFolderService oDFolderService = new ODFolderService(_GraphServiceWrapper, _userId);

				if(_OneDriveToCreate.RootFilesToCreate != null)
				{
					oDFolderService.CreateFiles(_OneDriveToCreate.RootFilesToCreate, "", "Root");
				}

				if(_OneDriveToCreate.FoldersToCreate != null)
				{
					foreach (var folderToCreate in _OneDriveToCreate.FoldersToCreate)
					{
						oDFolderService.CreateFolders(folderToCreate, "");
					}
				}

				Logger.FileLogger.Info($"OneDrive '{_userId}' creation completed successfully.");
				Console.WriteLine($"OneDrive '{_userId}' creation completed successfully.");
			}
		}
	}
}
