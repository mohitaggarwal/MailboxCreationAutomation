using MailboxCreationAutomation;
using MailboxCreationAutomation.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomationConsole
{
	class Program
	{
		static void Main(string[] args)
		{
			try
			{
				string username = ConfigurationManager.AppSettings["username"];
				string password = ConfigurationManager.AppSettings["password"];
				string jsonFile = ConfigurationManager.AppSettings["jsonFile"];
				FoldersToCreate foldersToCreate = JsonConvert.DeserializeObject<FoldersToCreate>(File.ReadAllText(jsonFile));
				EWSServiceWrapper eWSServiceWrapper = new EWSServiceWrapper(username, password);
				CalendarEvents calendarEvents = new CalendarEvents(eWSServiceWrapper);
				calendarEvents.CreateEvent();
				//MailboxFolder mailboxFolder = new MailboxFolder(eWSServiceWrapper);
				//Console.WriteLine("Mails creation started ...");
				//mailboxFolder.CreateFolders(foldersToCreate);
				//Console.WriteLine("Mails created successfully.");
			}
			catch(Exception ex)
			{
				Console.WriteLine($"Exception occur while creating mails. Detail: {ex.Message}");
			}
		}
	}
}
