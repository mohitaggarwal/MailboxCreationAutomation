﻿using DPClientSDK;
using DPClientSDK.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using MailboxCreationAutomation.Model;
using MailboxCreationAutomation;

namespace TestConsoleApp
{
	class Program
	{
		static void RunIncrBackup(int number, DPClient dpClient, EWSServiceWrapper ewsServiceWrapper, string backupSpecName)
		{
			DPClientSDK.Logger.FileLogger.Info($"Preparing Incremental {number} data ...");
			Console.WriteLine($"Preparing Incremental {number} data.");

			MailboxToCreate mailboxToCreate = new MailboxData().GetIncrBackup(number);
			Mailbox mailbox = new Mailbox(ewsServiceWrapper, mailboxToCreate);
			mailbox.CreateMailbox();

			DPClientSDK.Logger.FileLogger.Info($"Preparing Incremental {number} data completed.");
			Console.WriteLine($"Preparing Incremental {number} data completed.");

			DPClientSDK.Logger.FileLogger.Info($"Running incremental {number} backup ...");
			Console.WriteLine($"Running incremental {number} backup ...");

			RunBackup(dpClient, backupSpecName, "incr");

			DPClientSDK.Logger.FileLogger.Info($"Incremental {number} backup completed successfully.");
			Console.WriteLine($"Incremental {number} backup completed successfully.");
		}

		static void RunFullBackup(DPClient dpClient, EWSServiceWrapper ewsServiceWrapper, string backupSpecName)
		{
			DPClientSDK.Logger.FileLogger.Info($"Preparing first full backup data.");
			Console.WriteLine($"Preparing first full backup data.");

			MailboxToCreate mailboxToCreate = new MailboxData().GetFirstFullBackup();
			Mailbox mailbox = new Mailbox(ewsServiceWrapper, mailboxToCreate);
			mailbox.CreateMailbox();

			DPClientSDK.Logger.FileLogger.Info($"Preparing first full backup data completed.");
			Console.WriteLine($"Preparing first full backup data completed.");

			DPClientSDK.Logger.FileLogger.Info($"Running first full backup ...");
			Console.WriteLine($"Running first full backup ...");

			RunBackup(dpClient, backupSpecName, "full");

			DPClientSDK.Logger.FileLogger.Info($"First full backup completed successfully.");
			Console.WriteLine($"First full backup completed successfully.");
		}

		static void PrepareData(int number, MailboxToCreate mailboxToCreate, Mailbox mailbox, string subject)
		{
			foreach (var mailsToCreate in mailboxToCreate.RootMailsToCreates)
			{
				mailsToCreate.Subject = subject;
			}

			foreach (var eventsToCreate in mailboxToCreate.RootCalendarEventsToCreates)
			{
				eventsToCreate.CalendarEvent.Subject = subject;
				eventsToCreate.CalendarEvent.Body = subject;
				eventsToCreate.CalendarEvent.StartDate = DateTime.Now.AddDays(number);
				eventsToCreate.CalendarEvent.EndDate = DateTime.Now.AddDays(number).AddHours(1);
			}
			mailbox.CreateMailbox();
		}

		static void RunBackup(DPClient dpClient, string backupSpecName, string mode)
		{
			RunBackupSpecResponse response = dpClient.RunBackup(backupSpecName, mode);
			bool isSessionActive = false;
			do
			{
				Thread.Sleep(5000);
				var activeSessions = dpClient.GetActiveSessions();
				if (activeSessions != null && activeSessions.Sessions != null && activeSessions.Sessions.Count > 0)
				{
					isSessionActive = activeSessions.Sessions.Any(
										x => x.SessionName.Equals(response.SessionId, StringComparison.OrdinalIgnoreCase));
				}
				else
					isSessionActive = false;
			} while (isSessionActive);
		}

		static void RunIncrBackup(int number, MailboxToCreate mailboxToCreate, Mailbox mailbox, DPClient dpClient)
		{

			DPClientSDK.Logger.FileLogger.Info($"Preparing Incremental {number} data ...");
			Console.WriteLine($"Preparing Incremental {number} data.");

			PrepareData(number, mailboxToCreate, mailbox, $"Testing Incr {number}");

			DPClientSDK.Logger.FileLogger.Info($"Preparing Incremental {number} data completed.");
			Console.WriteLine($"Preparing Incremental {number} data completed.");

			DPClientSDK.Logger.FileLogger.Info($"Running incremental {number} backup ...");
			Console.WriteLine($"Running incremental {number} backup ...");

			RunBackup(dpClient, "Incr_User_Spec", "incr");

			DPClientSDK.Logger.FileLogger.Info($"Incremental {number} backup completed successfully.");
			Console.WriteLine($"Incremental {number} backup completed successfully.");
		}

		static void RunFullBackup(MailboxToCreate mailboxToCreate, Mailbox mailbox, DPClient dpClient)
		{
			DPClientSDK.Logger.FileLogger.Info($"Preparing first full backup data.");
			Console.WriteLine($"Preparing first full backup data.");

			PrepareData(0, mailboxToCreate, mailbox, "Testing First Full Backup");

			DPClientSDK.Logger.FileLogger.Info($"Preparing first full backup data completed.");
			Console.WriteLine($"Preparing first full backup data completed.");

			DPClientSDK.Logger.FileLogger.Info($"Running first full backup ...");
			Console.WriteLine($"Running first full backup ...");

			RunBackup(dpClient, "Incr_User_Spec", "full");

			DPClientSDK.Logger.FileLogger.Info($"First full backup completed successfully.");
			Console.WriteLine($"First full backup completed successfully.");
		}

		static void Main(string[] args)
		{
			try
			{
				int incrToRun = 3;
				EWSServiceWrapper ewsServiceWrapper = null; 
				BasicAuthInfo basicAuthInfo = new BasicAuthInfo
				{
					Username = "LidiaH@dpo365backup.onmicrosoft.com",
					Password = "Novell@12345"
				};
				ewsServiceWrapper = new EWSServiceWrapper(basicAuthInfo);
				DPClient dpClient = new DPClient("administrator|iwf1118201|iwf1118201.hpeswlab.net", "Data*pr0", "iwf1118201.hpeswlab.net");

				// First Full backup
				RunFullBackup(dpClient, ewsServiceWrapper, "Lidia_Spec");

				for (int i = 1; i <= incrToRun; i++)
				{
					RunIncrBackup(i, dpClient, ewsServiceWrapper, "Lidia_Spec");
				}



				//string jsonFile = @"TestJson2.json";
				//JsonSerializerSettings settings = new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.Auto };
				//MailboxToCreate mailboxToCreate = JsonConvert.DeserializeObject<MailboxToCreate>(File.ReadAllText(jsonFile), settings);
				//Mailbox mailbox = new Mailbox(ewsServiceWrapper, mailboxToCreate);
				//// First Full backup
				//RunFullBackup(mailboxToCreate, mailbox, dpClient);

				//for(int i=1; i<=incrToRun; i++)
				//{
				//	RunIncrBackup(i, mailboxToCreate, mailbox, dpClient);
				//}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Error: {ex.Message}");
			}
		}
	}
}
