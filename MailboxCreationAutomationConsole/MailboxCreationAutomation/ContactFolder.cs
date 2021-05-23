using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class ContactFolder
	{
		EWSServiceWrapper _EWSServiceWrapper;
		Contacts _Contacts;

		public ContactFolder(EWSServiceWrapper eWSServiceWrapper)
		{
			_EWSServiceWrapper = eWSServiceWrapper;
			_Contacts = new Contacts(_EWSServiceWrapper);
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
			folderView.PropertySet = PropertySet.IdOnly;
			SearchFilter.IsEqualTo isEqualTo = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, displayName);
			var folderResults = _EWSServiceWrapper.ExecuteCall(() => _EWSServiceWrapper.ExchangeService.FindFolders(parentFolder, isEqualTo, folderView));
			return folderResults.FirstOrDefault();
		}

		public string CreateFolder(
			string displayName,
			string parentFolderId)
		{
			Folder existFolder = GetFolder(displayName, parentFolderId);
			if (existFolder == null)
			{
				ContactsFolder folder = new ContactsFolder(_EWSServiceWrapper.ExchangeService);
				folder.DisplayName = displayName;
				_EWSServiceWrapper.ExecuteCall(() => folder.Save(parentFolderId));
				Logger.FileLogger.Info($"Contact folder '{displayName}' and id '{folder.Id.UniqueId}' created successfully.");
				return folder.Id.UniqueId;
			}
			else
			{
				Logger.FileLogger.Info($"Contact folder '{displayName}' and id '{existFolder.Id.UniqueId}' already exists.");
				return existFolder.Id.UniqueId;
			}
		}

		public string CreateFolderAtRoot(
			string displayName)
		{
			Folder existFolder = GetFolder(displayName, WellKnownFolderName.MsgFolderRoot);
			if (existFolder == null)
			{
				ContactsFolder folder = new ContactsFolder(_EWSServiceWrapper.ExchangeService);
				folder.DisplayName = displayName;
				_EWSServiceWrapper.ExecuteCall(() => folder.Save(WellKnownFolderName.MsgFolderRoot));
				Logger.FileLogger.Info($"Contact folder '{displayName}' and id '{folder.Id.UniqueId}' created successfully.");
				return folder.Id.UniqueId;
			}
			else
			{
				Logger.FileLogger.Info($"Contact folder '{displayName}' and id '{existFolder.Id.UniqueId}' already exists.");
				return existFolder.Id.UniqueId;
			}
		}

		public void CreateNestedFolders(
			string prefix,
			string parentFolderId,
			int currentLevel,
			int toLevel,
			List<ContactsToCreate> contactsToCreateList)
		{
			if (currentLevel > 0 && currentLevel <= toLevel)
			{
				string folderDisplayName = $"{prefix}_Level_{currentLevel}";
				string folderId = CreateFolder(folderDisplayName, parentFolderId);
				CreateContacts(contactsToCreateList, folderId, folderDisplayName);

				CreateNestedFolders(prefix, folderId, ++currentLevel, toLevel, contactsToCreateList);
			}
		}

		public void CreateFolders(ContactFoldersToCreate contactFoldersToCreate, string rootFolderId)
		{
			if (contactFoldersToCreate != null)
			{
				for (int i = 1; i <= contactFoldersToCreate.Count; i++)
				{
					string folderName = $"{contactFoldersToCreate.Prefix}_{i}";
					string folderId = CreateFolder(folderName, rootFolderId);
					CreateContacts(contactFoldersToCreate.ContactsToCreateList, folderId, folderName);

					CreateNestedFolders(folderName, folderId, 1, contactFoldersToCreate.Levels, contactFoldersToCreate.ContactsToCreateList);
				}
			}
		}

		public void DeleteFolder(string folderId)
		{
			Folder folder = Folder.Bind(_EWSServiceWrapper.ExchangeService, folderId);
			folder.Delete(DeleteMode.HardDelete);
		}

		public void CreateContacts(List<ContactsToCreate> contactToCreateList, string folderId, string prefix)
		{
			if (contactToCreateList != null)
			{
				int i = 1;
				foreach (var calendarsToCreate in contactToCreateList)
				{
					_Contacts.CreateContacts(calendarsToCreate, folderId, $"{prefix}_Set_{i}");
					i++;
				}
			}
		}
	}
}
