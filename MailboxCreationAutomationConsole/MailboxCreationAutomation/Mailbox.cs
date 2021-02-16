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
		}

		public void CreateMailbox()
		{
			if (_MailboxToCreate != null && _EWSServiceWrapper != null)
			{
				Logger.Username = _EWSServiceWrapper.Username;
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
	}
}
