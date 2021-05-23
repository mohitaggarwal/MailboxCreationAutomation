using DPClientSDK.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace DPClientSDK
{
	public class DPClient
	{
		private string _Username;
		private string _Password;
		private HttpClientWrapper _HttpClientWrapper;

		public TokenInformation Token { get; set; }

		public DPClient(string username, string password, string cellManager)
		{
			_Username = username;
			_Password = password;
			_HttpClientWrapper = new HttpClientWrapper();
			_HttpClientWrapper.BaseUri = new Uri(string.Format(DPConstants.BaseAddress, cellManager));
		}

		private Exception GetWebException(Exception ex)
		{
			if (ex is WebException || ex == null)
				return ex;
			else
				return GetWebException(ex.InnerException);
		}

		private void GetInnerException(Exception ex, ref string message)
		{
			if (!string.IsNullOrEmpty(ex.Message))
			{
				message += $"{ex.Message}\n";
			}
			if (ex.InnerException != null)
				GetInnerException(ex.InnerException, ref message);
		}

		public void GenerateToken()
		{
			try
			{
				TokenRequest tokenRequest = new TokenRequest
				{
					Username = _Username,
					Password = _Password,
					ClientId = "dp-gui",
					GrantType = "password"
				};

				string formData = $"username={_Username}&password={_Password}&client_id=dp-gui&grant_type=password";
				Logger.FileLogger.Info($"Token request: {formData}");
				string response = _HttpClientWrapper.Post("auth/realms/DataProtector/protocol/openid-connect/token", formData, "application/x-www-form-urlencoded");
				Logger.FileLogger.Info($"Token response: {response}");
				Token = JsonConvert.DeserializeObject<TokenInformation>(response);
				Token.TokenExpiresIn = DateTime.Now.AddSeconds(Token.ExpiresIn);
				_HttpClientWrapper.AccessToken = Token.AccessToken;
			}
			catch(AggregateException ex)
			{
				var webExcep = GetWebException(ex);
				if(webExcep != null)
				{
					Logger.FileLogger.Error($"Unable to generate token. Detail: {ex.Message}");
					throw;
				}
				else
				{
					string message = string.Empty;
					GetInnerException(ex, ref message);
					Logger.FileLogger.Error($"Unable to generate token. Detail: {message}");
					throw;

				}
			}
			catch(WebException ex)
			{
				Logger.FileLogger.Error($"Unable to generate token. Detail: {ex.Message}");
				throw;
			}
		}

		public void ValidateToken()
		{
			if (Token != null)
			{
				int variance = (Token.ExpiresIn * 10) / 100;
				variance = variance > 60 ? 60 : variance;
				if (Token.TokenExpiresIn.AddSeconds(-variance) < DateTime.Now)
				{
					GenerateToken();
				}
			}
			else
				GenerateToken();
		}

		public void GetBackupSpecification()
		{
			ValidateToken();
			try
			{
				string response = _HttpClientWrapper.Get("/dp-gui/dp-scheduler-gui/restws/backupspec");
				Logger.FileLogger.Info($"Backup Spec response: {response}");
			}
			catch (AggregateException ex)
			{
				var webExcep = GetWebException(ex);
				if (webExcep != null)
				{
					Logger.FileLogger.Error($"Unable to get backup specification. Detail: {ex.Message}");
					throw;
				}
				else
				{
					string message = string.Empty;
					GetInnerException(ex, ref message);
					Logger.FileLogger.Error($"Unable to get backup specification. Detail: {message}");
					throw;

				}
			}
			catch (WebException ex)
			{
				Logger.FileLogger.Error($"Unable to get backup specification. Detail: {ex.Message}");
				throw;
			}
		}

		public RunBackupSpecResponse RunBackup(string specName, string mode)
		{
			ValidateToken();
			try
			{
				RunBackupSpecRequest runBackupSpecRequest = new RunBackupSpecRequest
				{
					SpecificationName = specName,
					Mode = mode,
					ApplicationType = "M365",
					Load = "high",
					Monitor = "show"
				};
				string json = JsonConvert.SerializeObject(runBackupSpecRequest);
				Logger.FileLogger.Info($"Run backup json: {json}");
				string response = _HttpClientWrapper.Post("dp-protection/restws/v1/backup/m365/exchange", json);
				Logger.FileLogger.Info($"Run backup response: {response}");
				return JsonConvert.DeserializeObject<RunBackupSpecResponse>(response);
			}
			catch (AggregateException ex)
			{
				var webExcep = GetWebException(ex);

				if (webExcep != null)
				{
					Logger.FileLogger.Error($"Unable to run backup specification. Detail: {ex.Message}");
					throw;
				}
				else
				{
					string message = string.Empty;
					GetInnerException(ex, ref message);
					Logger.FileLogger.Error($"Unable to run backup specification. Detail: {message}");
					throw;

				}
			}
			catch (WebException ex)
			{
				Logger.FileLogger.Error($"Unable to run backup specification. Detail: {ex.Message}");
				throw;
			}
		}

		public ActiveSessions GetActiveSessions()
		{
			ValidateToken();
			try
			{
				string response = _HttpClientWrapper.Get("/dp-protection/restws/monitor/v1/sessions/active?pagesize=100&pagenum=1");
				Logger.FileLogger.Info($"Get active sessions response: {response}");
				return JsonConvert.DeserializeObject<ActiveSessions>(response);
			}
			catch (AggregateException ex)
			{
				var webExcep = GetWebException(ex);

				if (webExcep != null)
				{
					Logger.FileLogger.Error($"Unable to get active sessions. Detail: {ex.Message}");
					throw;
				}
				else
				{
					string message = string.Empty;
					GetInnerException(ex, ref message);
					Logger.FileLogger.Error($"Unable to get active sessions. Detail: {message}");
					throw;

				}
			}
			catch (WebException ex)
			{
				Logger.FileLogger.Error($"Unable to get active sessions. Detail: {ex.Message}");
				throw;
			}
		}

		public ActiveSessions GetActiveSessions(string json)
		{
			return JsonConvert.DeserializeObject<ActiveSessions>(json);
		}
	}
}