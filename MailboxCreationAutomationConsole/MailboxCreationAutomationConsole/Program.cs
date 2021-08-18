using MailboxCreationAutomation;
using MailboxCreationAutomation.Model;
using NDesk.Options;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxCreationAutomationConsole
{
	class Program
	{
		static string GetInnerExceptionMessage(Exception ex, string message)
		{
			if (ex == null)
				return message;
			else
				return GetInnerExceptionMessage(ex.InnerException, $"{message}\n{ex.Message}");
		}

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
				{ "authenticationType=", "Type of authentication: 'Basic' or 'OAuth'.",
				  v => commandLineOptions.AuthenticationType = v },
				{ "username=", "Mailbox username." ,
				  v => commandLineOptions.Username = v },
				{ "password=", "Mailbox password." ,
				  v => commandLineOptions.Password = v },
				{ "clientId=", "Azure app client Id." ,
				  v => commandLineOptions.ClientId = v },
				{ "clientSecret=", "Azure app client secret." ,
				  v => commandLineOptions.ClientSecret = v },
				{ "tenantId=", "Azure app tenant Id." ,
				  v => commandLineOptions.TenantId = v },
				{ "impersonateUser=", "Impersonate username." ,
				  v => commandLineOptions.ImpersonateUser = v },
				{ "jsonFile=", "Json file to create mailbox." ,
				  v => commandLineOptions.JsonFile = v },
				{ "createUser", "Create user" ,
				  v => commandLineOptions.CreateUser = v != null },
				{ "deleteUser=", "User to delete." ,
				  v => commandLineOptions.DeleteUser = v },
				{ "createOneDrive", "Create onedrive" ,
				  v => commandLineOptions.CreateOneDrive = v != null },
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
				JsonFile = ConfigurationManager.AppSettings["jsonFile"]
			};
			return commandLineOptions;
		}

		static void GraphTest()
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

			OAuthInfo oAuthInfo = new OAuthInfo
			{
				ClientId = "acc98107-38dc-4d18-8f9e-b341f643bd07",
				ClientSecret = "3RYhAubhqCJstuUZ/VvUw5YD9VQ//5m4kXxiEVWy6Kg=",
				TenantId = "dpo365backup.onmicrosoft.com"
			};

			GraphServiceWrapper graphServiceWrapper = new GraphServiceWrapper(oAuthInfo);
			MailboxUser mailboxUser = new MailboxUser(graphServiceWrapper);
			//mailboxUser.GetUserId("throttlingUser@dpo365backup.onmicrosoft.com");
			string userId = mailboxUser.CreateUser(userToCreate);
			mailboxUser.AssignLicense(userId, userToCreate.LicenseSkuId);
			mailboxUser.DeleteUser(userId);
		}

		public static void Run(string[] args)
		{
			try
			{
				CommandLineOptions commandLineOptions = args.Count() > 0 ? ParseCommandLine(args) : GetFromAppConfig();

				if (!commandLineOptions.ShowHelp)
				{
					BasicAuthInfo basicAuthInfo = null;
					OAuthInfo oAuthInfo = null;
					if (commandLineOptions.AuthenticationType.Equals("Basic", StringComparison.OrdinalIgnoreCase))
					{
						basicAuthInfo = new BasicAuthInfo
						{
							Username = commandLineOptions.Username,
							Password = commandLineOptions.Password
						};
					}
					else
					{
						oAuthInfo = new OAuthInfo
						{
							ClientId = commandLineOptions.ClientId,
							ClientSecret = commandLineOptions.ClientSecret,
							TenantId = commandLineOptions.TenantId,
							ImpersonateUser = commandLineOptions.ImpersonateUser
						};

					}

					if(commandLineOptions.CreateOneDrive)
					{
						GraphServiceWrapper graphServiceWrapper = new GraphServiceWrapper(oAuthInfo);
						JsonSerializerSettings settings = new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.Auto };
						OneDriveToCreate oneDriveToCreate = JsonConvert.DeserializeObject<OneDriveToCreate>(File.ReadAllText(commandLineOptions.JsonFile), settings);
						OneDrive oneDrive = new OneDrive(graphServiceWrapper, oneDriveToCreate, oAuthInfo.ImpersonateUser);
						oneDrive.CreateOneDrive();
					}
					else if (commandLineOptions.CreateUser)
					{
						GraphServiceWrapper graphServiceWrapper = new GraphServiceWrapper(oAuthInfo);
						JsonSerializerSettings settings = new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.Auto };
						UserToCreate userToCreate = JsonConvert.DeserializeObject<UserToCreate>(File.ReadAllText(commandLineOptions.JsonFile), settings);
						MailboxUser mailboxUser = new MailboxUser(graphServiceWrapper);
						string userId = mailboxUser.CreateUser(userToCreate);
						mailboxUser.AssignLicense(userId, userToCreate.LicenseSkuId);
					}
					else if (!string.IsNullOrEmpty(commandLineOptions.DeleteUser))
					{
						GraphServiceWrapper graphServiceWrapper = new GraphServiceWrapper(oAuthInfo);
						MailboxUser mailboxUser = new MailboxUser(graphServiceWrapper);
						mailboxUser.DeleteUser(commandLineOptions.DeleteUser);
					}
					else
					{
						EWSServiceWrapper ewsServiceWrapper = oAuthInfo == null ? new EWSServiceWrapper(basicAuthInfo) : new EWSServiceWrapper(oAuthInfo);
						JsonSerializerSettings settings = new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.Auto };
						MailboxToCreate mailboxToCreate = JsonConvert.DeserializeObject<MailboxToCreate>(File.ReadAllText(commandLineOptions.JsonFile), settings);
						Mailbox mailbox = new Mailbox(ewsServiceWrapper, mailboxToCreate);
						mailbox.CreateMailbox();
					}
				}
			}
			catch (Exception ex)
			{
				Logger.FileLogger.Error($"Exception occur while creating a mailbox. Detail: {GetInnerExceptionMessage(ex, ex.Message)}");
				Console.WriteLine($"Exception occur while creating a mailbox. Detail: {GetInnerExceptionMessage(ex, ex.Message)}");
			}
		}

		public static void TestCheckSum()
		{
			byte[] bytes1 = Encoding.ASCII.GetBytes("abc");
			byte[] bytes2 = Encoding.ASCII.GetBytes("def");
			byte[] bytes3 = Encoding.ASCII.GetBytes("abcdeff");
			QuickXorHash quickXorHash1 = new QuickXorHash();
			quickXorHash1.TransformBlock(bytes1, 0, bytes1.Length, bytes1, 0);
			quickXorHash1.TransformBlock(bytes2, 0, bytes2.Length, bytes2, 0);
			quickXorHash1.TransformFinalBlock(new byte[0], 0, 0);
			string hashInChunks = Convert.ToBase64String(quickXorHash1.Hash);

			QuickXorHash quickXorHash2 = new QuickXorHash();
			quickXorHash2.TransformBlock(bytes3, 0, bytes3.Length, bytes3, 0);
			quickXorHash2.TransformFinalBlock(new byte[0], 0, 0);
			string hashFull = Convert.ToBase64String(quickXorHash2.Hash);

			Console.WriteLine($"HashInChunks: {hashInChunks}, HashFull: {hashFull}, IsEqual: {hashFull.Equals(hashInChunks)}");
		}

		public static string CalculateHash(string filePath)
		{
			QuickXorHash quickXorHash = new QuickXorHash();
			using (var stream = File.OpenRead(filePath))
			{
				byte[] buffer = new byte[5 * 1024 * 1024];
				int bytesRead;
				while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
				{
					quickXorHash.TransformBlock(buffer, 0, buffer.Length, buffer, 0);
					Console.WriteLine($"Bytes Read: {bytesRead}");
				}
				quickXorHash.TransformFinalBlock(new byte[0], 0, 0);
				return Convert.ToBase64String(quickXorHash.Hash);
			}
		}

		static void Main(string[] args)
		{
			Run(args);
		}
	}
}
