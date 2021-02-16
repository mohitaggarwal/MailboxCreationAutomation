using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class CalendarEventsToCreate
	{
		public CalendarEventToCreate CalendarEvent { get; set; }
		public int Count { get; set; }
		public List<AttachmentsToCreate> AttachmentsToCreateList { get; set; }
	}

	public class CalendarEventToCreate
	{
		public string Subject { get; set; }
		public string Body { get; set; }
		public DateTime StartDate { get; set; }
		public DateTime EndDate { get; set; }
		public string Location { get; set; }
		public int ReminderMinutesBeforeStart { get; set; }
		public bool IsAllDayEvent { get; set; }
		public List<string> RequiredAttendees { get; set; }
		public List<string> OptionalAttendees { get; set; }
		public RecurrenceToCreate RecurrenceToCreate { get; set; }
	}

	public class RecurrenceToCreate
	{
		public int? NumberOfOccurrences { get; set; }
		public DateTime? EndDate { get; set; }
	}

	public class DailyPatternToCreate : RecurrenceToCreate
	{
		public int Interval { get; set; }
	}

	public class MonthlyPatternToCreate : RecurrenceToCreate
	{
		public int Interval { get; set; }
		public int DayOfMonth { get; set; }
	}

	public class RelativeMonthlyPatternToCreate : RecurrenceToCreate
	{
		public int Interval { get; set; }
		public int DayOfTheWeekIndex { get; set; }
		public int DayOfTheWeek { get; set; }
	}

	public class WeeklyPatternToCreate : RecurrenceToCreate
	{
		public int Interval { get; set; }
		public List<int> DayOfTheWeek { get; set; }
	}

	public class YearlyPatternToCreate : RecurrenceToCreate
	{
		public int DayOfMonth { get; set; }
		public int Month { get; set; }
	}

	public class RelativeYearlyPatternToCreate : RecurrenceToCreate
	{
		public int DayOfTheWeekIndex { get; set; }
		public int DayOfTheWeek { get; set; }
		public int Month { get; set; }
	}

}
