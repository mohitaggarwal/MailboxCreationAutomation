using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class MailboxToCreate
	{
		public string RootFolder { get; set; }
		public List<MailsToCreate> RootMailsToCreates { get; set; }
		public List<FoldersToCreate> FoldersToCreateList { get; set; }
		public string RootCalendar { get; set; }
		public List<CalendarEventsToCreate> RootCalendarEventsToCreates { get; set; }
		public List<CalendarsToCreate> CalendarsToCreateList { get; set; }
		public string RootContactFolder { get; set; }
		public List<ContactsToCreate> RootContactsToCreates { get; set; }
		public List<ContactFoldersToCreate> ContactFoldersToCreateList { get; set; }
	}
}
