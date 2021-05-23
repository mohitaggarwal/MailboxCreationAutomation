using MailboxCreationAutomation.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsoleApp
{
	public class MailboxData
	{
		public ContactsToCreate GetContactsToCreate()
		{
			return new ContactsToCreate
			{
				Count = 3,
				ContactToCreate = new ContactToCreate
				{
					GivenName = Constants.FirstFullBackupContactName,
					Surname = string.Empty,
					CompanyName = "Micro Focus",
					BussinessPhone = "+91-9986385033",
					HomeAddress = new AddressToCreate
					{
						Street = "27-J Yaseen Road",
						City = "Amritsar",
						State = "Punjab",
						CountryOrRegion = "India",
						PostalCode = "143001"
					},
					EmailAddress = Constants.FirstFullBackupContactEmail
				}
			};
		}

		public MailsToCreate GetMailsToCreate()
		{
			return new MailsToCreate
			{
				Count = 3,
				To = new List<string>
								{
									Constants.FirstFullBackupToRecipient
								},
				From = Constants.FirstFullBackupFromRecipient,
				Subject = Constants.FirstFullBackupMailSubject,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
								{
									new AttachmentsToCreate
									{
										Count = 1,
										AttachmentSizeInKB = 100
									}
								}
			};
		}

		public CalendarEventsToCreate GetDailyPatternEvent()
		{
			return new CalendarEventsToCreate
			{
				Count = 2,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
				{
					new AttachmentsToCreate
					{
						Count = 1,
						AttachmentSizeInKB = 30
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = Constants.FirstFullBackupEventSubject,
					Body = Constants.FirstFullBackupEventBody,
					StartDate = DateTime.Now,
					EndDate = DateTime.Now.AddHours(1),
					Location = "Meet Room 1",
					ReminderMinutesBeforeStart = 10,
					IsAllDayEvent = false,
					RequiredAttendees = new List<string>
					{
						"MultipleRestore1User@dpo365backup.onmicrosoft.com",
						"MultipleRestore3User@dpo365backup.onmicrosoft.com"
					},
					OptionalAttendees = new List<string>
					{
						"TemplateMailbox@dpo365backup.onmicrosoft.com"
					},
					RecurrenceToCreate = new DailyPatternToCreate
					{
						Interval = 1,
						NumberOfOccurrences = 10
					}
				}
			};
		}

		public CalendarEventsToCreate GetWeeklyPatternEvent()
		{
			return new CalendarEventsToCreate
			{
				Count = 2,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
				{
					new AttachmentsToCreate
					{
						Count = 1,
						AttachmentSizeInKB = 30
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = Constants.FirstFullBackupEventSubject,
					Body = Constants.FirstFullBackupEventBody,
					StartDate = DateTime.Now,
					EndDate = DateTime.Now.AddHours(1),
					Location = "Meet Room 1",
					ReminderMinutesBeforeStart = 10,
					IsAllDayEvent = false,
					RequiredAttendees = new List<string>
					{
						"MultipleRestore1User@dpo365backup.onmicrosoft.com",
						"MultipleRestore3User@dpo365backup.onmicrosoft.com"
					},
					OptionalAttendees = new List<string>
					{
						"TemplateMailbox@dpo365backup.onmicrosoft.com"
					},
					RecurrenceToCreate = new WeeklyPatternToCreate
					{
						Interval = 1,
						DayOfTheWeek = new List<int>
								{
									1,
									2
								},
						NumberOfOccurrences = 10
					}
				}
			};
		}

		public CalendarEventsToCreate GetMonthlyPatternEvent()
		{
			return new CalendarEventsToCreate
			{
				Count = 2,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
				{
					new AttachmentsToCreate
					{
						Count = 1,
						AttachmentSizeInKB = 30
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = Constants.FirstFullBackupEventSubject,
					Body = Constants.FirstFullBackupEventBody,
					StartDate = DateTime.Now,
					EndDate = DateTime.Now.AddHours(1),
					Location = "Meet Room 1",
					ReminderMinutesBeforeStart = 10,
					IsAllDayEvent = false,
					RequiredAttendees = new List<string>
					{
						"MultipleRestore1User@dpo365backup.onmicrosoft.com",
						"MultipleRestore3User@dpo365backup.onmicrosoft.com"
					},
					OptionalAttendees = new List<string>
					{
						"TemplateMailbox@dpo365backup.onmicrosoft.com"
					},
					RecurrenceToCreate = new MonthlyPatternToCreate
					{
						Interval = 1,
						DayOfMonth = 4,
						NumberOfOccurrences = 10
					}
				}
			};
		}

		public CalendarEventsToCreate GetYearlyPatternEvent()
		{
			return new CalendarEventsToCreate
			{
				Count = 2,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
				{
					new AttachmentsToCreate
					{
						Count = 1,
						AttachmentSizeInKB = 30
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = Constants.FirstFullBackupEventSubject,
					Body = Constants.FirstFullBackupEventBody,
					StartDate = DateTime.Now,
					EndDate = DateTime.Now.AddHours(1),
					Location = "Meet Room 1",
					ReminderMinutesBeforeStart = 10,
					IsAllDayEvent = false,
					RequiredAttendees = new List<string>
					{
						"MultipleRestore1User@dpo365backup.onmicrosoft.com",
						"MultipleRestore3User@dpo365backup.onmicrosoft.com"
					},
					OptionalAttendees = new List<string>
					{
						"TemplateMailbox@dpo365backup.onmicrosoft.com"
					},
					RecurrenceToCreate = new YearlyPatternToCreate
					{
						DayOfMonth = 24,
						Month = 5,
						NumberOfOccurrences = 10
					}
				}
			};
		}

		public MailboxToCreate GetFirstFullBackup()
		{
			MailboxToCreate mailboxToCreate = new MailboxToCreate
			{
				RootFolder = Constants.FirstFullBackupRootFolder,
				RootMailsToCreates = new List<MailsToCreate>
				{
					GetMailsToCreate()
				},
				RootCalendar = "Calendar",
				RootCalendarEventsToCreates = new List<CalendarEventsToCreate>
				{
					GetDailyPatternEvent(),
					GetYearlyPatternEvent()
				},
				RootContactFolder = "Contacts",
				RootContactsToCreates = new List<ContactsToCreate>
				{
					GetContactsToCreate()
				}
			};
			return mailboxToCreate;
		}


		public ContactsToCreate GetContactsToCreate(int n)
		{
			return new ContactsToCreate
			{
				Count = 3,
				ContactToCreate = new ContactToCreate
				{
					GivenName = string.Format(Constants.IncrBackupContactName, n),
					Surname = string.Empty,
					CompanyName = "Micro Focus",
					BussinessPhone = "+91-9986385033",
					HomeAddress = new AddressToCreate
					{
						Street = "27-J Yaseen Road",
						City = "Amritsar",
						State = "Punjab",
						CountryOrRegion = "India",
						PostalCode = "143001"
					},
					EmailAddress = string.Format(Constants.IncrBackupContactEmail, n)
				}
			};
		}

		public MailsToCreate GetMailsToCreate(int n)
		{
			return new MailsToCreate
			{
				Count = 3,
				To = new List<string>
								{
									Constants.FirstFullBackupToRecipient
								},
				From = Constants.FirstFullBackupFromRecipient,
				Subject = string.Format(Constants.IncrBackupMailSubject, n),
				AttachmentsToCreateList = new List<AttachmentsToCreate>
								{
									new AttachmentsToCreate
									{
										Count = 1,
										AttachmentSizeInKB = 100
									}
								}
			};
		}

		public CalendarEventsToCreate GetDailyPatternEvent(int n)
		{
			return new CalendarEventsToCreate
			{
				Count = 2,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
				{
					new AttachmentsToCreate
					{
						Count = 1,
						AttachmentSizeInKB = 30
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = string.Format(Constants.IncrBackupEventSubject, n),
					Body = string.Format(Constants.IncrBackupEventBody, n),
					StartDate = DateTime.Now.AddDays(n).AddHours(n),
					EndDate = DateTime.Now.AddDays(n).AddHours(n+1),
					Location = "Meet Room 1",
					ReminderMinutesBeforeStart = 10,
					IsAllDayEvent = false,
					RequiredAttendees = new List<string>
					{
						"MultipleRestore1User@dpo365backup.onmicrosoft.com",
						"MultipleRestore3User@dpo365backup.onmicrosoft.com"
					},
					OptionalAttendees = new List<string>
					{
						"TemplateMailbox@dpo365backup.onmicrosoft.com"
					},
					RecurrenceToCreate = new DailyPatternToCreate
					{
						Interval = 1,
						NumberOfOccurrences = 10
					}
				}
			};
		}

		public CalendarEventsToCreate GetWeeklyPatternEvent(int n)
		{
			return new CalendarEventsToCreate
			{
				Count = 2,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
				{
					new AttachmentsToCreate
					{
						Count = 1,
						AttachmentSizeInKB = 30
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = string.Format(Constants.IncrBackupEventSubject, n),
					Body = string.Format(Constants.IncrBackupEventBody, n),
					StartDate = DateTime.Now,
					EndDate = DateTime.Now.AddHours(1),
					Location = "Meet Room 1",
					ReminderMinutesBeforeStart = 10,
					IsAllDayEvent = false,
					RequiredAttendees = new List<string>
					{
						"MultipleRestore1User@dpo365backup.onmicrosoft.com",
						"MultipleRestore3User@dpo365backup.onmicrosoft.com"
					},
					OptionalAttendees = new List<string>
					{
						"TemplateMailbox@dpo365backup.onmicrosoft.com"
					},
					RecurrenceToCreate = new WeeklyPatternToCreate
					{
						Interval = 1,
						DayOfTheWeek = new List<int>
								{
									1,
									2
								},
						NumberOfOccurrences = 10
					}
				}
			};
		}

		public CalendarEventsToCreate GetMonthlyPatternEvent(int n)
		{
			return new CalendarEventsToCreate
			{
				Count = 2,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
				{
					new AttachmentsToCreate
					{
						Count = 1,
						AttachmentSizeInKB = 30
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = string.Format(Constants.IncrBackupEventSubject, n),
					Body = string.Format(Constants.IncrBackupEventBody, n),
					StartDate = DateTime.Now,
					EndDate = DateTime.Now.AddHours(1),
					Location = "Meet Room 1",
					ReminderMinutesBeforeStart = 10,
					IsAllDayEvent = false,
					RequiredAttendees = new List<string>
					{
						"MultipleRestore1User@dpo365backup.onmicrosoft.com",
						"MultipleRestore3User@dpo365backup.onmicrosoft.com"
					},
					OptionalAttendees = new List<string>
					{
						"TemplateMailbox@dpo365backup.onmicrosoft.com"
					},
					RecurrenceToCreate = new MonthlyPatternToCreate
					{
						Interval = 1,
						DayOfMonth = 4,
						NumberOfOccurrences = 10
					}
				}
			};
		}

		public CalendarEventsToCreate GetYearlyPatternEvent(int n)
		{
			return new CalendarEventsToCreate
			{
				Count = 2,
				AttachmentsToCreateList = new List<AttachmentsToCreate>
				{
					new AttachmentsToCreate
					{
						Count = 1,
						AttachmentSizeInKB = 30
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = string.Format(Constants.IncrBackupEventSubject, n),
					Body = string.Format(Constants.IncrBackupEventBody, n),
					StartDate = DateTime.Now,
					EndDate = DateTime.Now.AddHours(1),
					Location = "Meet Room 1",
					ReminderMinutesBeforeStart = 10,
					IsAllDayEvent = false,
					RequiredAttendees = new List<string>
					{
						"MultipleRestore1User@dpo365backup.onmicrosoft.com",
						"MultipleRestore3User@dpo365backup.onmicrosoft.com"
					},
					OptionalAttendees = new List<string>
					{
						"TemplateMailbox@dpo365backup.onmicrosoft.com"
					},
					RecurrenceToCreate = new YearlyPatternToCreate
					{
						DayOfMonth = 24,
						Month = 5,
						NumberOfOccurrences = 10
					}
				}
			};
		}

		public MailboxToCreate GetIncrBackup(int n)
		{
			MailboxToCreate mailboxToCreate = new MailboxToCreate
			{
				RootFolder = string.Format(Constants.IncrBackupRootFolder, n),
				RootMailsToCreates = new List<MailsToCreate>
				{
					GetMailsToCreate(n)
				},
				RootCalendar = "Calendar",
				RootCalendarEventsToCreates = new List<CalendarEventsToCreate>
				{
					GetDailyPatternEvent(n)
				},
				RootContactFolder = "Contacts",
				RootContactsToCreates = new List<ContactsToCreate>
				{
					GetContactsToCreate(n)
				}
			};
			return mailboxToCreate;
		}
	}
}
