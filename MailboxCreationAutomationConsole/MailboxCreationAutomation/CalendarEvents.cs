using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

        private void GetDateTime(CalendarEventToCreate calendarEvent, int count, out DateTime startDateTime, out DateTime endDateTime)
		{
            if (typeof(WeeklyPatternToCreate) == calendarEvent.RecurrenceToCreate?.GetType())
            {
                int days = 7 * (count / 24);
                startDateTime = calendarEvent.StartDate.AddDays(days).AddHours(count % 24);
                endDateTime = calendarEvent.EndDate.AddDays(days).AddHours(count % 24);
            }
            else if (typeof(DailyPatternToCreate) == calendarEvent.RecurrenceToCreate?.GetType())
            {
                startDateTime = calendarEvent.StartDate.AddDays(count / 24).AddHours(count % 24);
                endDateTime = calendarEvent.EndDate.AddDays(count / 24).AddHours(count % 24);
            }
            else if (typeof(MonthlyPatternToCreate) == calendarEvent.RecurrenceToCreate?.GetType() ||
                    typeof(RelativeMonthlyPatternToCreate) == calendarEvent.RecurrenceToCreate?.GetType())
            {
                int days = 30 * (count / 24);
                startDateTime = calendarEvent.StartDate.AddDays(days).AddHours(count % 24);
                endDateTime = calendarEvent.EndDate.AddDays(days).AddHours(count % 24);
            }
            else if (typeof(YearlyPatternToCreate) == calendarEvent.RecurrenceToCreate?.GetType() ||
                    typeof(RelativeYearlyPatternToCreate) == calendarEvent.RecurrenceToCreate?.GetType())
            {
                int days = 365 * (count / 24);
                startDateTime = calendarEvent.StartDate.AddDays(days).AddHours(count % 24);
                endDateTime = calendarEvent.EndDate.AddDays(days).AddHours(count % 24);
            }
			else
			{
                startDateTime = calendarEvent.StartDate.AddDays(count / 24).AddHours(count % 24);
                endDateTime = calendarEvent.EndDate.AddDays(count / 24).AddHours(count % 24);
			}
        }

        private Recurrence GetRecurrence(WeeklyPatternToCreate weeklyPatternToCreate, DateTime startDateTime)
		{
            if (weeklyPatternToCreate != null)
            {
                DayOfTheWeek[] dow = null;
                if (weeklyPatternToCreate.DayOfTheWeek.Count > 0)
                {
                    dow = new DayOfTheWeek[weeklyPatternToCreate.DayOfTheWeek.Count];
                    int i = 0;
                    foreach (var dayOfTheWeek in weeklyPatternToCreate.DayOfTheWeek)
                    {
                        dow[i] = (DayOfTheWeek)dayOfTheWeek;
                        i++;
                    }
                }
                Recurrence recurrence = new Recurrence.WeeklyPattern(startDateTime, weeklyPatternToCreate.Interval, dow);
                recurrence.StartDate = startDateTime;
                recurrence.NumberOfOccurrences = weeklyPatternToCreate.NumberOfOccurrences;
                if (weeklyPatternToCreate.EndDate != null)
                {
                    recurrence.EndDate = weeklyPatternToCreate.EndDate;
                }
                return recurrence;
            }
            else
                return null;
        }

        private Recurrence GetRecurrence(DailyPatternToCreate dailyPatternToCreate, DateTime startDateTime)
        {
            if (dailyPatternToCreate != null)
            {
                Recurrence recurrence = new Recurrence.DailyPattern(startDateTime, dailyPatternToCreate.Interval);
                recurrence.NumberOfOccurrences = dailyPatternToCreate.NumberOfOccurrences;
                if (dailyPatternToCreate.EndDate != null)
                    recurrence.EndDate = dailyPatternToCreate.EndDate;
                return recurrence;
            }
            else
                return null;
        }

        private Recurrence GetRecurrence(MonthlyPatternToCreate monthlyPatternToCreate, DateTime startDateTime)
        {
            if (monthlyPatternToCreate != null)
            {
                Recurrence recurrence = new Recurrence.MonthlyPattern(startDateTime, monthlyPatternToCreate.Interval, monthlyPatternToCreate.DayOfMonth);
                recurrence.NumberOfOccurrences = monthlyPatternToCreate.NumberOfOccurrences;
                if (monthlyPatternToCreate.EndDate != null)
                    recurrence.EndDate = monthlyPatternToCreate.EndDate;
                return recurrence;
            }
            else
                return null;
        }

        private Recurrence GetRecurrence(RelativeMonthlyPatternToCreate relativeMonthlyPatternToCreate, DateTime startDateTime)
        {
            if (relativeMonthlyPatternToCreate != null)
            {
                Recurrence recurrence = new Recurrence.RelativeMonthlyPattern(
                                                startDateTime, 
                                                relativeMonthlyPatternToCreate.Interval,
                                                (DayOfTheWeek)relativeMonthlyPatternToCreate.DayOfTheWeek,
                                                (DayOfTheWeekIndex)relativeMonthlyPatternToCreate.DayOfTheWeekIndex);
                recurrence.NumberOfOccurrences = relativeMonthlyPatternToCreate.NumberOfOccurrences;
                if (relativeMonthlyPatternToCreate.EndDate != null)
                    recurrence.EndDate = relativeMonthlyPatternToCreate.EndDate;
                return recurrence;
            }
            else
                return null;
        }

        private Recurrence GetRecurrence(YearlyPatternToCreate yearlyPatternToCreate, DateTime startDateTime)
        {
            if (yearlyPatternToCreate != null)
            {
                Recurrence recurrence = new Recurrence.YearlyPattern(startDateTime, (Month)yearlyPatternToCreate.Month, yearlyPatternToCreate.DayOfMonth);
                recurrence.NumberOfOccurrences = yearlyPatternToCreate.NumberOfOccurrences;
                if (yearlyPatternToCreate.EndDate != null)
                    recurrence.EndDate = yearlyPatternToCreate.EndDate;
                return recurrence;
            }
            else
                return null;
        }

        private Recurrence GetRecurrence(RelativeYearlyPatternToCreate relativeYearlyPatternToCreate, DateTime startDateTime)
        {
            if (relativeYearlyPatternToCreate != null)
            {
                Recurrence recurrence = new Recurrence.RelativeYearlyPattern(
                                                startDateTime, 
                                                (Month)relativeYearlyPatternToCreate.Month, 
                                                (DayOfTheWeek)relativeYearlyPatternToCreate.DayOfTheWeek,
                                                (DayOfTheWeekIndex)relativeYearlyPatternToCreate.DayOfTheWeekIndex);
                recurrence.NumberOfOccurrences = relativeYearlyPatternToCreate.NumberOfOccurrences;
                if (relativeYearlyPatternToCreate.EndDate != null)
                    recurrence.EndDate = relativeYearlyPatternToCreate.EndDate;
                return recurrence;
            }
            else
                return null;
        }

        private Appointment GetAppointment(CalendarEventToCreate calendarEvent, string prefix, int count)
		{
            if (calendarEvent != null)
            {
                Appointment recurrMeeting = new Appointment(_EWSServiceWrapper.ExchangeService);
                DateTime startDateTime;
                DateTime endDateTime;
                GetDateTime(calendarEvent, count, out startDateTime, out endDateTime);
                // Set the properties you need to create a meeting.
                recurrMeeting.Subject = $"{calendarEvent.Subject}_Event_{count}_{prefix}";
                recurrMeeting.Body = $"{calendarEvent.Body}_Event_{count}";
                recurrMeeting.Start = startDateTime;
                recurrMeeting.End = endDateTime;
                recurrMeeting.Location = calendarEvent.Location;
                recurrMeeting.IsAllDayEvent = calendarEvent.IsAllDayEvent;
                recurrMeeting.ReminderMinutesBeforeStart = calendarEvent.ReminderMinutesBeforeStart;

                if (calendarEvent.RequiredAttendees != null)
                {
                    foreach (var requiredAttendee in calendarEvent.RequiredAttendees)
                    {
                        recurrMeeting.RequiredAttendees.Add(requiredAttendee);
                    }
                }

                if (calendarEvent.OptionalAttendees != null)
                {
                    foreach (var optionalAttendee in calendarEvent.OptionalAttendees)
                    {
                        recurrMeeting.OptionalAttendees.Add(optionalAttendee);
                    }
                }

                if (calendarEvent.RecurrenceToCreate != null)
                {
                    if (typeof(WeeklyPatternToCreate) == calendarEvent.RecurrenceToCreate.GetType())
                    {
                        recurrMeeting.Recurrence = GetRecurrence(calendarEvent.RecurrenceToCreate as WeeklyPatternToCreate, startDateTime);
                    }
                    else if (typeof(DailyPatternToCreate) == calendarEvent.RecurrenceToCreate.GetType())
                    {
                        recurrMeeting.Recurrence = GetRecurrence(calendarEvent.RecurrenceToCreate as DailyPatternToCreate, startDateTime);
                    }
                    else if (typeof(MonthlyPatternToCreate) == calendarEvent.RecurrenceToCreate.GetType())
                    {
                        recurrMeeting.Recurrence = GetRecurrence(calendarEvent.RecurrenceToCreate as MonthlyPatternToCreate, startDateTime);
                    }
                    else if (typeof(YearlyPatternToCreate) == calendarEvent.RecurrenceToCreate.GetType())
                    {
                        recurrMeeting.Recurrence = GetRecurrence(calendarEvent.RecurrenceToCreate as YearlyPatternToCreate, startDateTime);
                    }
                    else if (typeof(RelativeMonthlyPatternToCreate) == calendarEvent.RecurrenceToCreate.GetType())
                    {
                        recurrMeeting.Recurrence = GetRecurrence(calendarEvent.RecurrenceToCreate as RelativeMonthlyPatternToCreate, startDateTime);
                    }
                    else if (typeof(RelativeYearlyPatternToCreate) == calendarEvent.RecurrenceToCreate.GetType())
                    {
                        recurrMeeting.Recurrence = GetRecurrence(calendarEvent.RecurrenceToCreate as RelativeYearlyPatternToCreate, startDateTime);
                    }
                }
                return recurrMeeting;
            }
            else
                return null;
        }

        private void CreateAttachment(string fileName, int sizeInKB)
        {
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                fs.SetLength(sizeInKB * 1024);
            }
        }

        private void EventAttachments(string directory, AttachmentsToCreate attachmentsToCreate, string prefix, ref Appointment appointment)
        {
            if (attachmentsToCreate != null)
            {
                for (int i = 1; i <= attachmentsToCreate.Count; i++)
                {
                    string attachmentName = $"{prefix}_{i}_sizeInKB_{attachmentsToCreate.AttachmentSizeInKB}.txt";
                    string fileName = Path.Combine(directory, attachmentName);
                    CreateAttachment(fileName, attachmentsToCreate.AttachmentSizeInKB);
                    // Add a file attachment by using a stream, and specify the name of the attachment.
                    // The email attachment is named FourthAttachment.txt.
                    FileStream theStream = new FileStream(fileName, FileMode.OpenOrCreate);
                    // In this example, theStream is a Stream object that represents the content of the file to attach.
                    appointment.Attachments.AddFileAttachment(attachmentName, theStream);
                }
            }
        }

        public void CreateEvents(CalendarEventsToCreate calendarEventsToCreate, string folderId, string prefix)
		{

            for (int i = 1; i <= calendarEventsToCreate.Count; i++)
            {
                bool needRetry = true;
                int retryCount = 0;
                do
                {
                    //string directory = Path.Combine(System.Environment.
                    //								 GetFolderPath(
                    //									 Environment.SpecialFolder.CommonApplicationData),
                    //								"TempAttachmentFolder");
                    string directory = $"{_EWSServiceWrapper.Username}_{DateTime.Now.ToString("yyyy - MM - dd HH - mm - ss - fff")}";
                    DirectoryInfo di = Directory.CreateDirectory(directory);
                    try
                    {
                        Appointment appointment = GetAppointment(calendarEventsToCreate.CalendarEvent, prefix, i);
                        if (calendarEventsToCreate.AttachmentsToCreateList != null)
                        {
                            int j = 1;
                            foreach (var attachmentToCreate in calendarEventsToCreate.AttachmentsToCreateList)
                            {
                                EventAttachments(directory, attachmentToCreate, $"{prefix}_Set_{j}", ref appointment);
                                j++;
                            }
                        }
                        appointment.Save(folderId, SendInvitationsMode.SendToNone);
                        Logger.FileLogger.Info($"Event with subject '{appointment.Subject}' and id '{appointment.Id.UniqueId}' created successfully.");
                        needRetry = false;
                    }
                    catch (ServerBusyException ex)
                    {
                        Console.WriteLine($"Server is busy. Retrying after {ex.BackOffMilliseconds / 1000}sec");
                        Logger.FileLogger.Warning($"Server is busy. Retrying after {ex.BackOffMilliseconds / 1000}sec");
                        Thread.Sleep(ex.BackOffMilliseconds);
                        needRetry = true;
                        retryCount++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Exception occur while creating events. Detail: {ex.Message}");
                        Logger.FileLogger.Error($"Exception occur while creating events. Detail: {ex.Message}");
                        Console.WriteLine($"Retrying after {EWSServiceConstants.RETRY_AFTER / 1000}sec");
                        Logger.FileLogger.Warning($"Retrying after {EWSServiceConstants.RETRY_AFTER / 1000}sec");
                        Thread.Sleep(EWSServiceConstants.RETRY_AFTER);
                        needRetry = true;
                        retryCount++;
                    }
                    finally
                    {
                        if(di.Exists)
                            di.Delete(true);
                    }
                } while (needRetry && retryCount < EWSServiceConstants.RETRY_COUNT);
            }
		}
	}
}
