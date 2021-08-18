using MailboxCreationAutomation.Model;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class MailboxUser
	{
		private GraphServiceWrapper _GraphServiceWrapper;

		public MailboxUser(GraphServiceWrapper graphServiceWrapper)
		{
			_GraphServiceWrapper = graphServiceWrapper;
		}

		public string CreateUser(UserToCreate userToCreate)
		{
			Logger.FileLogger.Info($"User '{userToCreate.UserPrinicpleName}' creation started ...");
			Console.WriteLine($"User '{userToCreate.UserPrinicpleName}' creation started ...");
			var user = new User
			{
				AccountEnabled = true,
				DisplayName = userToCreate.DisplayName,
				MailNickname = userToCreate.MailNickName,
				UserPrincipalName = userToCreate.UserPrinicpleName,
				PasswordProfile = new PasswordProfile
				{
					ForceChangePasswordNextSignIn = false,
					Password = userToCreate.Password
				},
				UsageLocation = userToCreate.UsageLocation
			};

			var userCreated = _GraphServiceWrapper.GraphServiceClient.Users
								.Request()
								.AddAsync(user).Result;
			Logger.FileLogger.Info($"User '{userToCreate.UserPrinicpleName}' creation completed successfully.");
			Console.WriteLine($"User '{userToCreate.UserPrinicpleName}' creation completed successfully.");
			return userCreated.Id;
		}

		public void AssignLicense(string userId, string licenseSkuId)
		{
			Logger.FileLogger.Info($"Assigning license to user '{userId}' ...");
			Console.WriteLine($"Assigning license to user '{userId}' ...");
			var addLicenses = new List<AssignedLicense>
				{
					new AssignedLicense
					{
						SkuId = Guid.Parse(licenseSkuId)
					}
				};
			_GraphServiceWrapper.GraphServiceClient.Users[userId]
				.AssignLicense(addLicenses, new List<Guid>())
				.Request()
				.PostAsync().Wait();
			Logger.FileLogger.Info($"Assigning license to user '{userId}' completed successfully.");
			Console.WriteLine($"Assigning license to user '{userId}' completed successfully.");
		}

		public string GetLicenseSkuid(string userId)
		{
			var licenseDetails = _GraphServiceWrapper.GraphServiceClient.Users[userId].LicenseDetails
										.Request()
										.GetAsync().Result;
			return licenseDetails.FirstOrDefault().SkuId.ToString();
		}

		public string GetUserId(string userpricipleName)
		{
			var user = _GraphServiceWrapper.GraphServiceClient.Users[userpricipleName]
										.Request()
										.GetAsync().Result;
			return user.Id;
		}

		public void DeleteUser(string userId)
		{
			Logger.FileLogger.Info($"Deleting user '{userId}' ...");
			Console.WriteLine($"Deleting user '{userId}' ...");
			_GraphServiceWrapper.GraphServiceClient.Users[userId].Request().DeleteAsync().Wait();
			Logger.FileLogger.Info($"Deleting user '{userId}' completed successfully.");
			Console.WriteLine($"Deleting user '{userId}' completed successfully.");
		}
	}
}
