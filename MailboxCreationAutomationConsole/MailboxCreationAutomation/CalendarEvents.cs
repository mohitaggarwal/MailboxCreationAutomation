using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class CalendarEvents
	{
		EWSServiceWrapper _EWSServiceWrapper;

		public CalendarEvents(EWSServiceWrapper eWSServiceWrapper)
		{
			_EWSServiceWrapper = eWSServiceWrapper;
		}

		public void CreateEvent()
		{
            Appointment recurrMeeting = new Appointment(_EWSServiceWrapper.ExchangeService);
            // Set the properties you need to create a meeting.
            recurrMeeting.Subject = "Weekly Update Meeting";
            recurrMeeting.Body = "Come hear about how the project is coming along!";
            recurrMeeting.Start = DateTime.Now.AddDays(1);
            recurrMeeting.End = recurrMeeting.Start.AddHours(1);
            recurrMeeting.Location = "Contoso Main Gallery";
            recurrMeeting.RequiredAttendees.Add("Mack@contoso.com");
            recurrMeeting.RequiredAttendees.Add("Sadie@contoso.com");
            recurrMeeting.RequiredAttendees.Add("Magdalena@contoso.com"); recurrMeeting.ReminderMinutesBeforeStart = 30;

            DayOfTheWeek[] dow = new DayOfTheWeek[] { (DayOfTheWeek)recurrMeeting.Start.DayOfWeek };
            // The following are the recurrence-specific properties for the meeting.
            recurrMeeting.Recurrence = new Recurrence.WeeklyPattern(recurrMeeting.Start.Date, 1, dow);
            recurrMeeting.Recurrence.StartDate = recurrMeeting.Start.Date;
            recurrMeeting.Recurrence.NumberOfOccurrences = 10;
            // This method results in in a CreateItem call to EWS.
            recurrMeeting.Save(SendInvitationsMode.SendToNone);
        }
	}
}
