using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class MailboxFolder
	{
		EWSServiceWrapper _EWSServiceWrapper;
		MailboxMails _MailboxMails;

		public MailboxFolder(EWSServiceWrapper eWSServiceWrapper)
		{
			_EWSServiceWrapper = eWSServiceWrapper;
			_MailboxMails = new MailboxMails(_EWSServiceWrapper);
		}

		public Folder GetFolder(WellKnownFolderName wellKnownFolderName)
		{
			return _EWSServiceWrapper.ExecuteCall(() => Folder.Bind(_EWSServiceWrapper.ExchangeService, 
				wellKnownFolderName, PropertySet.FirstClassProperties));
		}

		public Folder GetFolder(string displayName, string parentFolderId)
		{
			FolderView folderView = new FolderView(1);
			folderView.Traversal = FolderTraversal.Shallow;
			folderView.PropertySet = PropertySet.IdOnly;
			SearchFilter.IsEqualTo isEqualTo = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, displayName);
			var folderResults = _EWSServiceWrapper.ExecuteCall(() => _EWSServiceWrapper.ExchangeService.FindFolders(parentFolderId, isEqualTo, folderView));
			return folderResults.FirstOrDefault();
		}

		public Folder GetFolder(string displayName, WellKnownFolderName parentFolder)
		{
			FolderView folderView = new FolderView(1);
			folderView.Traversal = FolderTraversal.Shallow;
			folderView.PropertySet = new PropertySet() { FolderSchema.Id, FolderSchema.TotalCount };
			SearchFilter.IsEqualTo isEqualTo = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, displayName);
			var folderResults = _EWSServiceWrapper.ExecuteCall(() =>  _EWSServiceWrapper.ExchangeService.FindFolders(parentFolder, isEqualTo, folderView));
			return folderResults.FirstOrDefault();
		}

		public List<Folder> GetFolders(string parentFolderId, FolderTraversal folderTraversal)
		{
			List<Folder> folders = new List<Folder>();
			bool moreAvail = false;
			do
			{
				FolderView folderView = new FolderView(512);
				folderView.Traversal = folderTraversal;
				folderView.PropertySet = new PropertySet() { FolderSchema.Id, FolderSchema.DisplayName, FolderSchema.TotalCount};
				var folderResults = _EWSServiceWrapper.ExecuteCall(() => _EWSServiceWrapper.ExchangeService.FindFolders(parentFolderId, folderView));
				folders.AddRange(folderResults.ToList());
				moreAvail = folderResults.MoreAvailable;
			} while (moreAvail);
			return folders;
		}

		public List<Folder> GetFolders(WellKnownFolderName parentFolder, FolderTraversal folderTraversal)
		{
			List<Folder> folders = new List<Folder>();
			bool moreAvail = false;
			do
			{
				FolderView folderView = new FolderView(512);
				folderView.Traversal = folderTraversal;
				folderView.PropertySet = new PropertySet() { FolderSchema.Id, FolderSchema.DisplayName, FolderSchema.TotalCount };
				var folderResults = _EWSServiceWrapper.ExecuteCall(() => _EWSServiceWrapper.ExchangeService.FindFolders(parentFolder, folderView));
				folders.AddRange(folderResults.ToList());
				moreAvail = folderResults.MoreAvailable;
			} while (moreAvail);
			return folders;
		}

		public List<Folder> GetFolders()
		{
			List<Folder> folders = new List<Folder>();
			PropertySet propertySet = new PropertySet
			{
				FolderSchema.Id,
				FolderSchema.DisplayName,
				FolderSchema.TotalCount,
				FolderSchema.FolderClass,
				FolderSchema.ParentFolderId
			};
			bool isMoreAvailable = false;
			string syncState = string.Empty;
			do
			{
				var folderChanges = _EWSServiceWrapper.ExecuteCall(() => 
											_EWSServiceWrapper.ExchangeService.SyncFolderHierarchy(
												WellKnownFolderName.MsgFolderRoot, propertySet, syncState));
				syncState = folderChanges.SyncState;
				isMoreAvailable = folderChanges.MoreChangesAvailable;
				folders.AddRange(folderChanges.Select(x => x.Folder));
			} while (isMoreAvailable);
			return folders;
		}

		public string CreateFolder(
			string displayName,
			string parentFolderId)
		{
			Folder existFolder = GetFolder(displayName, parentFolderId);
			if (existFolder == null)
			{
				Folder folder = new Folder(_EWSServiceWrapper.ExchangeService);
				folder.FolderClass = "IPF.Note";
				folder.DisplayName = displayName;
				_EWSServiceWrapper.ExecuteCall(() => folder.Save(parentFolderId));
				Logger.FileLogger.Info($"Folder '{displayName}' and id '{folder.Id.UniqueId}' created successfully.");
				return folder.Id.UniqueId;
			}
			else
			{
				Logger.FileLogger.Info($"Folder '{displayName}' and id '{existFolder.Id.UniqueId}' already exists.");
				return existFolder.Id.UniqueId;
			}
		}

		public string CreateFolderAtRoot(
			string displayName)
		{
			Folder existFolder = GetFolder(displayName, WellKnownFolderName.MsgFolderRoot);
			if (existFolder == null)
			{
				Folder folder = new Folder(_EWSServiceWrapper.ExchangeService);
				folder.FolderClass = "IPF.Note";
				folder.DisplayName = displayName;
				_EWSServiceWrapper.ExecuteCall(() => folder.Save(WellKnownFolderName.MsgFolderRoot));
				Logger.FileLogger.Info($"Folder '{displayName}' and id '{folder.Id.UniqueId}' created successfully.");
				return folder.Id.UniqueId;
			}
			else
			{
				Logger.FileLogger.Info($"Folder '{displayName}' and id '{existFolder.Id.UniqueId}' already exists.");
				return existFolder.Id.UniqueId;
			}
		}

		public void CreateNestedFolders(
			string prefix,
			string parentFolderId,
			int currentLevel,
			int toLevel,
			List<MailsToCreate> mailsToCreateList)
		{
			if (currentLevel > 0 && currentLevel <= toLevel)
			{
				string folderDisplayName = $"{prefix}_Level_{currentLevel}";
				string folderId = CreateFolder(folderDisplayName, parentFolderId);
				CreateMails(mailsToCreateList, folderId, folderDisplayName);

				CreateNestedFolders(prefix, folderId, ++currentLevel, toLevel, mailsToCreateList);
			}
		}

		public void CreateFolders(FoldersToCreate foldersToCreate, string rootFolderId)
		{
			if (foldersToCreate != null)
			{
				for (int i = 1; i <= foldersToCreate.Count; i++)
				{
					string folderName = $"{foldersToCreate.Prefix}_{i}";
					string folderId = CreateFolder(folderName, rootFolderId);
					CreateMails(foldersToCreate.MailsToCreateList, folderId, folderName);

					CreateNestedFolders(folderName, folderId, 1, foldersToCreate.Levels, foldersToCreate.MailsToCreateList);
				}
			}
		}

		public void DeleteFolder(string folderId)
		{
			Folder folder = Folder.Bind(_EWSServiceWrapper.ExchangeService, folderId);
			folder.Delete(DeleteMode.HardDelete);
		}

		public void CreateMails(List<MailsToCreate> mailsToCreateList, string folderId, string prefix)
		{
			if (mailsToCreateList != null)
			{
				int i = 1;
				foreach (var mailsToCreate in mailsToCreateList)
				{
					_MailboxMails.CreateMails(mailsToCreate, folderId, $"{prefix}_Mail_Set_{i}");
					i++;
				}
			}
		}

		public void PrintFolder(string displayName)
		{
			Folder rootFolder = GetFolder(displayName, WellKnownFolderName.MsgFolderRoot);
			if (rootFolder != null) {
				int mailCount = rootFolder.TotalCount;
				List<Folder> folders = GetFolders(rootFolder.Id.UniqueId, FolderTraversal.Deep);
				foreach (Folder folder in folders)
				{
					mailCount += folder.TotalCount;
					Console.WriteLine($"Folder: {folder.DisplayName}, count: {folder.TotalCount}");
				}
				Console.WriteLine($"Root Folder: {displayName}, Total count: {mailCount}");
			}
		}

		public void PrintMailboxSize()
		{
			long mailboxSize = 0;
			int mailCount = 0;
			var folders = GetFolders();
			foreach(var folder in folders)
			{
				long folderSize = 0;
				foreach(var folderItems in _MailboxMails.GetItems(folder.Id.UniqueId))
				{
					folderSize += folderItems.Size;
				}
				mailCount += folder.TotalCount;
				Logger.FileLogger.Info($"Folder: {folder.DisplayName}, count: {folder.TotalCount}, size: {folderSize}");
				Console.WriteLine($"Folder: {folder.DisplayName}, count: {folder.TotalCount}, size: {folderSize}");
				mailboxSize += folderSize;
			}
			Logger.FileLogger.Info($"Mailbox size: {mailboxSize}, Total mail count: {mailCount}");
			Console.WriteLine($"Mailbox size: {mailboxSize}");
		}

		public void EmptyMailbox()
		{
			List<Folder> folders = GetFolders(WellKnownFolderName.MsgFolderRoot, FolderTraversal.Shallow);
			foreach(var folder in folders)
			{
				Console.WriteLine($"Display Name: {folder.DisplayName}");
				if (folder is CalendarFolder)
				{
					Console.WriteLine($"Calendar folder.");
				}
				else
				{
					folder.Empty(DeleteMode.HardDelete, true);
				}
			}
		}
	}
}
