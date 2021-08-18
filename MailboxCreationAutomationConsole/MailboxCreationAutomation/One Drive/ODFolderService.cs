using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class ODFolderService
	{
		private GraphServiceWrapper _GraphServiceWrapper;
		private string _userId;
		private DriveItemService _DriveItemService;

		public ODFolderService(GraphServiceWrapper graphServiceWrapper, string userId)
		{
			_GraphServiceWrapper = graphServiceWrapper;
			_userId = userId;
			_DriveItemService = new DriveItemService(_GraphServiceWrapper, _userId);
		}

		private string GetInnerExceptionMessage(Exception ex, string message)
		{
			if (ex == null)
				return message;
			else
				return GetInnerExceptionMessage(ex.InnerException, $"{message}\n{ex.Message}");
		}

		public string CreateFolder(string name, string odFilePath)
		{
			var driveItem = new DriveItem
			{
				Name = name,
				Folder = new Folder
				{
				},
				AdditionalData = new Dictionary<string, object>()
				{
					{"@microsoft.graph.conflictBehavior", "rename"}
				}
			};
			if (string.IsNullOrEmpty(odFilePath))
			{
				var response = _GraphServiceWrapper.GraphServiceClient
							.Users[_userId]
							.Drive
							.Root
							.Children
							.Request()
							.AddAsync(driveItem)
							.Result;
				return response.Id;
			}
			else
			{
				var response = _GraphServiceWrapper.GraphServiceClient
							.Users[_userId]
							.Drive
							.Root
							.ItemWithPath(odFilePath)
							.Children
							.Request()
							.AddAsync(driveItem)
							.Result;
				return response.Id;
			}
		}

		public void CreateNestedFolders(
			string prefix,
			string odFilePath,
			int currentLevel,
			int toLevel,
			List<ODFileToCreate>filesToCreate)
		{
			if (currentLevel > 0 && currentLevel <= toLevel)
			{
				string folderDisplayName = $"{prefix}_Level_{currentLevel}";
				string folderId = CreateFolder(folderDisplayName, odFilePath);
				string odFilePathWithName = Path.Combine(odFilePath, folderDisplayName);
				CreateFiles(filesToCreate, odFilePathWithName, folderDisplayName);

				CreateNestedFolders(prefix, odFilePathWithName, ++currentLevel, toLevel, filesToCreate);
			}
		}

		public void CreateFolders(ODFolderToCreate foldersToCreate, string odFilePath)
		{
			if (foldersToCreate != null)
			{
				for (int i = 1; i <= foldersToCreate.Count; i++)
				{
					string folderName = $"{foldersToCreate.Prefix}_{i}";
					string folderId = CreateFolder(folderName, odFilePath);
					string odFilePathWithName = Path.Combine(odFilePath, folderName);
					CreateFiles(foldersToCreate.FilesToCreate, odFilePathWithName, folderName);

					CreateNestedFolders(folderName, odFilePathWithName, 1, foldersToCreate.Levels, foldersToCreate.FilesToCreate);
				}
			}
		}

		public void CreateFiles(List<ODFileToCreate> filesToCreate, string odFilePath, string prefix)
		{
			if (filesToCreate != null)
			{
				int i = 1;
				foreach (var fileToCreate in filesToCreate)
				{
					_DriveItemService.CreatFiles(fileToCreate, odFilePath, $"{prefix}_Set_{i}");
					i++;
				}
			}
		}
	}
}
