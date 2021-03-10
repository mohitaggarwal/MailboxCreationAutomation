using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class Mailbox
	{
		private EWSServiceWrapper _EWSServiceWrapper;
		private MailboxToCreate _MailboxToCreate;

		public Mailbox(EWSServiceWrapper ewsServiceWrapper, MailboxToCreate mailboxToCreate)
		{
			_EWSServiceWrapper = ewsServiceWrapper;
			_MailboxToCreate = mailboxToCreate;
			Logger.Username = _EWSServiceWrapper.Username;
		}

		public Mailbox(EWSServiceWrapper ewsServiceWrapper)
		{
			_EWSServiceWrapper = ewsServiceWrapper;
			Logger.Username = _EWSServiceWrapper.Username;
		}

		private List<Folder> GetFilterFolders(IEnumerable<Folder> mailFolders, HashSet<string> ignoreFolderSet)
		{
			List<Folder> folders = new List<Folder>();
			foreach(var mailFolder in mailFolders)
			{
				if(ignoreFolderSet.Contains(mailFolder.ParentFolderId.UniqueId))
				{
					ignoreFolderSet.Add(mailFolder.Id.UniqueId);
				}
				else if(!ignoreFolderSet.Contains(mailFolder.Id.UniqueId))
				{
					folders.Add(mailFolder);
				}
			}
			return folders;
		}

		private void GetFolderItems(IEnumerable<Folder> folders)
		{
			MailboxMails mailboxMails = new MailboxMails(_EWSServiceWrapper);
			foreach (var folder in folders)
			{
				var folderItems = mailboxMails.GetItems(folder.Id.UniqueId);
				Logger.FileLogger.Info($"Folder {folder.DisplayName}, items: {folderItems.Count()}");
				Console.WriteLine($"Folder {folder.DisplayName}, items: {folderItems.Count()}");
				foreach (var folderItem in folderItems)
				{
					Logger.FileLogger.Info($"Getting item Id: {folderItem.Id.UniqueId}, Subject: {folderItem.Subject}");
					mailboxMails.ExportItem(folderItem.Id);
				}
			}
		}

		private void GetFolderItemsInBulk(IEnumerable<Folder> folders)
		{
			MailboxMails mailboxMails = new MailboxMails(_EWSServiceWrapper);
			foreach (var folder in folders)
			{
				var folderItems = mailboxMails.GetItems(folder.Id.UniqueId);
				Logger.FileLogger.Info($"Folder {folder.DisplayName}, items: {folderItems.Count()}");
				Console.WriteLine($"Folder {folder.DisplayName}, items: {folderItems.Count()}");
				int pageSize = 100;
				int elementsRead = 0;
				bool isMoreAvailable = false;
				do
				{
					var foldersPage = folderItems.Skip(elementsRead).Take(pageSize);
					isMoreAvailable = foldersPage.Any();
					if (isMoreAvailable)
					{
						Logger.FileLogger.Info($"Getting Items range: {elementsRead + 1}-{elementsRead + pageSize}");
						mailboxMails.ExportItem(foldersPage.Select(x => x.Id).ToList());
						elementsRead += pageSize;
					}
				} while (isMoreAvailable);
			}
		}

		public void CreateMailbox()
		{
			if (_MailboxToCreate != null && _EWSServiceWrapper != null)
			{
				Logger.FileLogger.Info($"Mailbox '{_EWSServiceWrapper.Username}' creation started ...");
				Console.WriteLine($"Mailbox '{_EWSServiceWrapper.Username}' creation started ...");


				MailboxFolder mailboxFolder = new MailboxFolder(_EWSServiceWrapper);
				if (!string.IsNullOrEmpty(_MailboxToCreate.RootFolder))
				{
					string rootFolderId = mailboxFolder.CreateFolderAtRoot(_MailboxToCreate.RootFolder);
					mailboxFolder.CreateMails(_MailboxToCreate.RootMailsToCreates, rootFolderId, _MailboxToCreate.RootFolder);
					if (_MailboxToCreate.FoldersToCreateList != null)
					{
						foreach (var foldersToCreate in _MailboxToCreate.FoldersToCreateList)
						{
							mailboxFolder.CreateFolders(foldersToCreate, rootFolderId);
						}
					}
				}

				Calendar calendar = new Calendar(_EWSServiceWrapper);
				if (!string.IsNullOrEmpty(_MailboxToCreate.RootCalendar))
				{
					string rootCalendarId = calendar.CreateFolderAtRoot(_MailboxToCreate.RootCalendar);
					calendar.CreateEvents(_MailboxToCreate.RootCalendarEventsToCreates, rootCalendarId, _MailboxToCreate.RootCalendar);
					if (_MailboxToCreate.CalendarsToCreateList != null)
					{
						foreach (var calendarsToCreate in _MailboxToCreate.CalendarsToCreateList)
						{
							calendar.CreateFolders(calendarsToCreate, rootCalendarId);
						}
					}
				}

				Logger.FileLogger.Info($"Mailbox '{_EWSServiceWrapper.Username}' creation completed successfully.");
				Console.WriteLine($"Mailbox '{_EWSServiceWrapper.Username}' creation completed successfully.");
			}
		}

		public void DeleteFolder(string displayName)
		{
			MailboxFolder mailboxFolder = new MailboxFolder(_EWSServiceWrapper);
			Folder folder = mailboxFolder.GetFolder(displayName, WellKnownFolderName.DeletedItems);
			if(folder != null)
			{
				mailboxFolder.DeleteFolder(folder.Id.UniqueId);
			}
		}

		public void DeleteFolderInRoot(string displayName)
		{
			MailboxFolder mailboxFolder = new MailboxFolder(_EWSServiceWrapper);
			Folder folder = mailboxFolder.GetFolder(displayName, WellKnownFolderName.MsgFolderRoot);
			if (folder != null)
			{
				mailboxFolder.DeleteFolder(folder.Id.UniqueId);
			}
		}

		public void BackupMailbox(string folderPrefixToIgnore)
		{
			if (_EWSServiceWrapper != null)
			{
				Logger.FileLogger.Info($"Mailbox '{_EWSServiceWrapper.Username}' backup started ...");
				Console.WriteLine($"Mailbox '{_EWSServiceWrapper.Username}' backup started ...");
				MailboxFolder mailboxFolder = new MailboxFolder(_EWSServiceWrapper);
				List<Folder> folders = mailboxFolder.GetFolders();
				if (folders != null)
				{
					var mailFolders = folders.Where(x => !string.IsNullOrEmpty(x.FolderClass)
															&& x.FolderClass.StartsWith("IPF.Note", StringComparison.OrdinalIgnoreCase));
					var prefixFolders = mailFolders.Where(x => !string.IsNullOrEmpty(x.FolderClass)
															&& x.DisplayName.StartsWith(folderPrefixToIgnore, StringComparison.OrdinalIgnoreCase))
										.Select(x => x.Id.UniqueId);
					var prefixFolderSet = new HashSet<string>(prefixFolders);
					var mailFolderAfterIgnoreFolders = GetFilterFolders(mailFolders, prefixFolderSet);

					Logger.FileLogger.Info($"Mail Folders: {mailFolderAfterIgnoreFolders.Count()}");
					Console.WriteLine($"Mail Folders: {mailFolderAfterIgnoreFolders.Count()}");
					GetFolderItemsInBulk(mailFolderAfterIgnoreFolders);

					var calFolders = folders.Where(x => x is CalendarFolder);
					Logger.FileLogger.Info($"Calendar Folders: {calFolders.Count()}");
					Console.WriteLine($"Calendar Folders: {calFolders.Count()}");
					GetFolderItemsInBulk(calFolders);

				}
				Logger.FileLogger.Info($"Mailbox '{_EWSServiceWrapper.Username}' backup completed successfully.");
				Console.WriteLine($"Mailbox '{_EWSServiceWrapper.Username}' backup completed successfully.");
			}
		}

		public void PrintFolderDetails(string displayName)
		{
			MailboxFolder mailboxFolder = new MailboxFolder(_EWSServiceWrapper);
			mailboxFolder.PrintFolder(displayName);
		}
	}
}
