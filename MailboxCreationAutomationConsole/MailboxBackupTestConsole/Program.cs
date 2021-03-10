using MailboxCreationAutomation;
using MailboxCreationAutomation.Model;
using MailboxCreationAutomationConsole;
using NDesk.Options;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxBackupTestConsole
{
	class Program
	{
		static void ShowHelp(OptionSet p)
		{
			Console.WriteLine("Options:");
			p.WriteOptionDescriptions(Console.Out);
		}

		static CommandLineOptions ParseCommandLine(string[] args)
		{
			CommandLineOptions commandLineOptions = new CommandLineOptions();

			var p = new OptionSet()
			{
				{ "at|authenticationType=", "Type of authentication: 'Basic' or 'OAuth'.",
				  v => commandLineOptions.AuthenticationType = v },
				{ "u|username=", "Mailbox username." ,
				  v => commandLineOptions.Username = v },
				{ "p|password=", "Mailbox password." ,
				  v => commandLineOptions.Password = v },
				{ "ci|clientId=", "Azure app client Id." ,
				  v => commandLineOptions.ClientId = v },
				{ "cs|clientSecret=", "Azure app client secret." ,
				  v => commandLineOptions.ClientSecret = v },
				{ "t|tenantId=", "Azure app tenant Id." ,
				  v => commandLineOptions.TenantId = v },
				{ "iu|impersonateUser=", "Impoersonate username." ,
				  v => commandLineOptions.ImpersonateUser = v },
				{ "fp|folderPrefixToIgnore=", "Folder prefix to ignore for backup." ,
				  v => commandLineOptions.FolerPrefixToIgnore = v },
				{ "h|help",  "List different options",
				  v => commandLineOptions.ShowHelp = v != null }
			};

			List<string> extra;
			try
			{
				extra = p.Parse(args);
			}
			catch (OptionException e)
			{
				Console.WriteLine($"Error while parsing the command line arguments. Error: {e.Message}");
				commandLineOptions.ShowHelp = true;
			}

			if (commandLineOptions.ShowHelp)
			{
				ShowHelp(p);
			}

			return commandLineOptions;
		}

		static CommandLineOptions GetFromAppConfig()
		{
			CommandLineOptions commandLineOptions = new CommandLineOptions
			{
				AuthenticationType = ConfigurationManager.AppSettings["authenticationType"],
				Username = ConfigurationManager.AppSettings["username"],
				Password = ConfigurationManager.AppSettings["password"],
				ClientId = ConfigurationManager.AppSettings["clientId"],
				ClientSecret = ConfigurationManager.AppSettings["clientSecret"],
				TenantId = ConfigurationManager.AppSettings["tenantId"],
				ImpersonateUser = ConfigurationManager.AppSettings["impersonateUser"],
				FolerPrefixToIgnore = ConfigurationManager.AppSettings["folderPrefixToIgnore"]
			};
			return commandLineOptions;
		}

		static void Main(string[] args)
		{
			try
			{
				CommandLineOptions commandLineOptions = args.Count() > 0 ? ParseCommandLine(args) : GetFromAppConfig();

				if (!commandLineOptions.ShowHelp)
				{
					EWSServiceWrapper ewsServiceWrapper = null;
					if (commandLineOptions.AuthenticationType.Equals("Basic", StringComparison.OrdinalIgnoreCase))
					{
						BasicAuthInfo basicAuthInfo = new BasicAuthInfo
						{
							Username = commandLineOptions.Username,
							Password = commandLineOptions.Password
						};
						ewsServiceWrapper = new EWSServiceWrapper(basicAuthInfo);
					}
					else
					{
						OAuthInfo oAuthInfo = new OAuthInfo
						{
							ClientId = commandLineOptions.ClientId,
							ClientSecret = commandLineOptions.ClientSecret,
							TenantId = commandLineOptions.TenantId,
							ImpersonateUser = commandLineOptions.ImpersonateUser
						};
						ewsServiceWrapper = new EWSServiceWrapper(oAuthInfo);
					}

					Mailbox mailbox = new Mailbox(ewsServiceWrapper);
					mailbox.BackupMailbox(commandLineOptions.FolerPrefixToIgnore);
				}
			}
			catch (Exception ex)
			{
				Logger.FileLogger.Error($"Exception occur while backing-up a mailbox. Detail: { ex.Message}");
				Console.WriteLine($"Exception occur while  backing-up a mailbox. Detail: { ex.Message}");
			}
		}
	}
}
