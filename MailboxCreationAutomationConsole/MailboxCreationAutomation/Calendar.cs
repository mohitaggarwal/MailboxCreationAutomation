using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class Calendar
	{
		EWSServiceWrapper _EWSServiceWrapper;
		CalendarEvents _CalendarEvents;

		public Calendar(EWSServiceWrapper eWSServiceWrapper)
		{
			_EWSServiceWrapper = eWSServiceWrapper;
			_CalendarEvents = new CalendarEvents(_EWSServiceWrapper);
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
				CalendarFolder folder = new CalendarFolder(_EWSServiceWrapper.ExchangeService);
				folder.DisplayName = displayName;
				_EWSServiceWrapper.ExecuteCall(() => folder.Save(parentFolderId));
				Logger.FileLogger.Info($"Calendar '{displayName}' and id '{folder.Id.UniqueId}' created successfully.");
				return folder.Id.UniqueId;
			}
			else
			{
				Logger.FileLogger.Info($"Calendar '{displayName}' and id '{existFolder.Id.UniqueId}' already exists.");
				return existFolder.Id.UniqueId;
			}
		}

		public string CreateFolderAtRoot(
			string displayName)
		{
			Folder existFolder = GetFolder(displayName, WellKnownFolderName.MsgFolderRoot);
			if (existFolder == null)
			{
				CalendarFolder folder = new CalendarFolder(_EWSServiceWrapper.ExchangeService);
				folder.DisplayName = displayName;
				_EWSServiceWrapper.ExecuteCall(() => folder.Save(WellKnownFolderName.MsgFolderRoot));
				Logger.FileLogger.Info($"Calendar '{displayName}' and id '{folder.Id.UniqueId}' created successfully.");
				return folder.Id.UniqueId;
			}
			else
			{
				Logger.FileLogger.Info($"Calendar '{displayName}' and id '{existFolder.Id.UniqueId}' already exists.");
				return existFolder.Id.UniqueId;
			}
		}

		public void CreateNestedFolders(
			string prefix,
			string parentFolderId,
			int currentLevel,
			int toLevel,
			List<CalendarEventsToCreate> calendarEventsToCreateList)
		{
			if (currentLevel > 0 && currentLevel <= toLevel)
			{
				string folderDisplayName = $"{prefix}_Level_{currentLevel}";
				string folderId = CreateFolder(folderDisplayName, parentFolderId);
				CreateEvents(calendarEventsToCreateList, folderId, folderDisplayName);

				CreateNestedFolders(prefix, folderId, ++currentLevel, toLevel, calendarEventsToCreateList);
			}
		}

		public void CreateFolders(CalendarsToCreate calendarsToCreate, string rootFolderId)
		{
			if (calendarsToCreate != null)
			{
				for (int i = 1; i <= calendarsToCreate.Count; i++)
				{
					string folderName = $"{calendarsToCreate.Prefix}_{i}";
					string folderId = CreateFolder(folderName, rootFolderId);
					CreateEvents(calendarsToCreate.CalendarEventsToCreateList, folderId, folderName);

					CreateNestedFolders(folderName, folderId, 1, calendarsToCreate.Levels, calendarsToCreate.CalendarEventsToCreateList);
				}
			}
		}

		public void DeleteFolder(string folderId)
		{
			Folder folder = Folder.Bind(_EWSServiceWrapper.ExchangeService, folderId);
			folder.Delete(DeleteMode.HardDelete);
		}

		public void CreateEvents(List<CalendarEventsToCreate> calendarEventsToCreateList, string folderId, string prefix)
		{
			if (calendarEventsToCreateList != null)
			{
				int i = 1;
				foreach (var calendarsToCreate in calendarEventsToCreateList)
				{
					_CalendarEvents.CreateEvents(calendarsToCreate, folderId, $"{prefix}_Event_Set_{i}");
					i++;
				}
			}
		}
	}
}
