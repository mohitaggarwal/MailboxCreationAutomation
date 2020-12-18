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
			return Folder.Bind(_EWSServiceWrapper.ExchangeService, 
				wellKnownFolderName, PropertySet.FirstClassProperties);
		}

		public string CreateFolder(
			string displayName,
			string parentFolderId)
		{
			Folder folder = new Folder(_EWSServiceWrapper.ExchangeService);
			folder.FolderClass = "IPF.Note";
			folder.DisplayName = displayName;
			_EWSServiceWrapper.ExecuteCall(() => folder.Save(parentFolderId));
			return folder.Id.UniqueId;
		}

		public string CreateFolderAtRoot(
			string displayName)
		{
			Folder folder = new Folder(_EWSServiceWrapper.ExchangeService);
			folder.FolderClass = "IPF.Note";
			folder.DisplayName = displayName;
			_EWSServiceWrapper.ExecuteCall(() => folder.Save(WellKnownFolderName.MsgFolderRoot));
			return folder.Id.UniqueId;
		}

		public void CreateNestedFolders(
			string displayName,
			string parentFolderId,
			int currentLevel,
			int toLevel,
			MailsToCreate mailsToCreate)
		{
			if (currentLevel > 0 && currentLevel <= toLevel)
			{
				string folderDisplayName = $"{displayName}_Level_{currentLevel}";
				string folderId = CreateFolder(folderDisplayName, parentFolderId);
				_MailboxMails.CreateMails(mailsToCreate, folderId);

				CreateNestedFolders(displayName, folderId, ++currentLevel, toLevel, mailsToCreate);
			}
		}

		public void CreateFolders(FoldersToCreate foldersToCreate)
		{
			string rootFolderId = CreateFolderAtRoot(foldersToCreate.RootFolder);
			for(int i = 1; i<= foldersToCreate.Folders; i++)
			{
				string folderName = $"{foldersToCreate.RootFolder}_{i}";
				string folderId = CreateFolder(folderName, rootFolderId);
				_MailboxMails.CreateMails(foldersToCreate.MailsToCreate, folderId);

				CreateNestedFolders(folderName, folderId, 1, foldersToCreate.Levels, foldersToCreate.MailsToCreate);
			}
		}

		public void DeleteFolder(string folderId)
		{
			Folder folder = Folder.Bind(_EWSServiceWrapper.ExchangeService, folderId);
			folder.Delete(DeleteMode.HardDelete);
		}
	}
}
