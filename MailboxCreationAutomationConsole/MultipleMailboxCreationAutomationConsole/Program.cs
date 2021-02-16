using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultipleMailboxCreationAutomationConsole
{
	class Program
	{
		static void GenerateFileWithMailboxes()
		{
			string text;
			for(int i = 1; i <= 50; i++)
			{
				if(i < 11)
				{
					text = $"DP-DevTestUser{i}@MFPublicCloud.onmicrosoft.com,Data*pr0,MailboxUser_10000.json";
				}
				else if( i > 10 && i < 26)
				{
					text = $"DP-DevTestUser{i}@MFPublicCloud.onmicrosoft.com,Data*pr0,MailboxUser_5000.json";
				}
				else
				{
					text = $"DP-DevTestUser{i}@MFPublicCloud.onmicrosoft.com,Data*pr0,MailboxUser_1000.json";
				}
				using (StreamWriter writer = new StreamWriter("mailboxesToCreate.txt", true, System.Text.Encoding.UTF8))
				{
					if (!string.IsNullOrEmpty(text))
					{
						writer.WriteLine(text);
					}
				}
			}
		}

		static void Main(string[] args)
		{
			// -authenticationType=Basic -username=mohit_admin@dpo365backup.onmicrosoft.com -password=Novell@1234 -jsonFile=D:\Workspace\Learning\MailboxCreationAutomation\MailboxCreationAutomationConsole\Output\TestJson.json
			try
			{
				if (args.Length == 1)
				{
					List<Tuple<string, string, string>> mailboxes = GetMailboxes(args[0]);
					foreach (var mailbox in mailboxes)
					{
						Process p = new Process();
						p.StartInfo = new ProcessStartInfo("MailboxCreationAutomationConsole.exe", $"-authenticationType=Basic -username={mailbox.Item1} -password={mailbox.Item2} -jsonFile={mailbox.Item3}");
						//p.StartInfo.WorkingDirectory = @"C:\Program Files\Chrome";
						p.StartInfo.CreateNoWindow = true;
						p.StartInfo.UseShellExecute = false;
						p.Start();
					}
				}
				string input = string.Empty;
				do
				{
					Console.WriteLine($"Type 'exit' to end the application");
					input = Console.ReadLine();
				} while (!input.Equals("exit", StringComparison.OrdinalIgnoreCase));
			}
			catch(Exception ex)
			{
				Console.WriteLine($"Error: {ex.Message}");
			}
		}

		static List<Tuple<string, string, string>> GetMailboxes(string filePath)
		{
			List<Tuple<string, string, string>> mailboxes = new List<Tuple<string, string, string>>();
			using (var reader = new StreamReader(filePath))
			{
				while (!reader.EndOfStream)
				{
					var line = reader.ReadLine();
					var values = line.Split(';', ',');

					if(values.Count() == 3)
					{
						mailboxes.Add(Tuple.Create(values[0], values[1], values[2]));
					}
				}
			}
			return mailboxes;
		}
	}
}
