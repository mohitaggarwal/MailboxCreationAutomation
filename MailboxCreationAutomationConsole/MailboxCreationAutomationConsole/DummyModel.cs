using MailboxCreationAutomation.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomationConsole
{
	public class DummyModel
	{
		public ContactsToCreate GetContactsToCreate()
		{
			return new ContactsToCreate
			{
				Count = 3,
				ContactToCreate = new ContactToCreate
				{
					GivenName = "Mohit",
					Surname = "Aggarwal",
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
					EmailAddress = "mohit.aggarwal2@microfocus.com"
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
									"mohit_admin@dpo365backup.onmicrosoft.com"
								},
				From = "Haripriya@hpdatapro.onmicrosoft.com",
				Subject = "Testing Mail for small Application",
				AttachmentsToCreateList = new List<AttachmentsToCreate>
								{
									new AttachmentsToCreate
									{
										Count = 3,
										AttachmentSizeInKB = 54
									},
									new AttachmentsToCreate
									{
										Count = 2,
										AttachmentSizeInKB = 124
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
					},
					new AttachmentsToCreate
					{
						Count = 5,
						AttachmentSizeInKB = 20
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = "Testing event subject DailyPatternToCreate",
					Body = "Testing event body DailyPatternToCreate",
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
					},
					new AttachmentsToCreate
					{
						Count = 5,
						AttachmentSizeInKB = 20
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = "Testing event subject WeeklyPatternToCreate",
					Body = "Testing event body WeeklyPatternToCreate",
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
					},
					new AttachmentsToCreate
					{
						Count = 5,
						AttachmentSizeInKB = 20
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = "Testing event subject MonthlyPatternToCreate",
					Body = "Testing event body MonthlyPatternToCreate",
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
					},
					new AttachmentsToCreate
					{
						Count = 5,
						AttachmentSizeInKB = 20
					}
				},
				CalendarEvent = new CalendarEventToCreate
				{
					Subject = "Testing event subject YearlyPatternToCreate",
					Body = "Testing event body YearlyPatternToCreate",
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

		public MailboxToCreate GetMailboxToCreate()
		{
			MailboxToCreate mailboxToCreate = new MailboxToCreate
			{
				RootFolder = "TestMailboxCreationAutomation2",
				FoldersToCreateList = new List<FoldersToCreate>
				{
					new FoldersToCreate
					{
						Prefix = "Test_Small_Mails",
						Count = 2,
						Levels = 1,
						MailsToCreateList = new List<MailsToCreate>
						{
							GetMailsToCreate()
						}
					}
				},
				RootCalendar = "Calendar",
				RootCalendarEventsToCreates = new List<CalendarEventsToCreate>
				{
					GetDailyPatternEvent(),
					GetYearlyPatternEvent()
				},
				CalendarsToCreateList = new List<CalendarsToCreate>
				{
					new CalendarsToCreate
					{

						Prefix = "Test_Small_Cal",
						Count = 2,
						Levels = 1,
						CalendarEventsToCreateList = new List<CalendarEventsToCreate>
						{
							GetWeeklyPatternEvent(),
							GetMonthlyPatternEvent()
						}
					}
				},
				RootContactFolder = "Contacts",
				ContactFoldersToCreateList = new List<ContactFoldersToCreate>
				{
					new ContactFoldersToCreate
					{
						Prefix = "Test_Contacts",
						Count = 2,
						Levels = 1,
						ContactsToCreateList = new List<ContactsToCreate>
						{
							GetContactsToCreate()
						}
					}
				}
			};
			return mailboxToCreate;
		}

		public string GetDummyJson()
		{
			string json = JsonConvert.SerializeObject(GetMailboxToCreate(), Formatting.Indented, new JsonSerializerSettings
							{
								TypeNameHandling = TypeNameHandling.Auto
							});
			return json;
		}

		public string GetDummyUserToCreate()
		{
			UserToCreate userToCreate = new UserToCreate
			{
				DisplayName = "Test user from Graph",
				MailNickName = "GraphUser",
				UserPrinicpleName = "testgraphuser@dpo365backup.onmicrosoft.com",
				Password = "Novell@1234",
				UsageLocation = "IN",
				LicenseSkuId = "c42b9cae-ea4f-4ab7-9717-81576235ccac"
			};

			string json = JsonConvert.SerializeObject(userToCreate, Formatting.Indented, new JsonSerializerSettings
			{
				TypeNameHandling = TypeNameHandling.Auto
			});
			return json;
		}
	}
}
