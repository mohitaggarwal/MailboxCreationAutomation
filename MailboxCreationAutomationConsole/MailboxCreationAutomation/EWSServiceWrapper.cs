using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Net;
using System.Threading;

namespace MailboxCreationAutomation
{
    public class EWSServiceWrapper
    {
        public ExchangeService ExchangeService { get; set; }
        public string Username { get; set; }

        public EWSServiceWrapper(BasicAuthInfo basicAuthInfo)
		{
            TraceListener traceListener = new TraceListener();
            ExchangeService = new ExchangeService(ExchangeVersion.Exchange2013);
            ExchangeService.Credentials = new WebCredentials(basicAuthInfo.Username, basicAuthInfo.Password);
            ExchangeService.Url = new Uri(EWSServiceConstants.EWS_URL);
            ExchangeService.Timeout = 500000;
            //ExchangeService.TraceListener = traceListener;
            //ExchangeService.TraceEnabled = true;
            ExchangeService.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
            Username = basicAuthInfo.Username;
        }

        public EWSServiceWrapper(OAuthInfo oAuthInfo)
        {
            TraceListener traceListener = new TraceListener();
            ExchangeService = new ExchangeService();
            ExchangeService.Url = new Uri(EWSServiceConstants.EWS_URL);
            ExchangeService.Credentials = new OAuthCredentials(GetOAuthToken(oAuthInfo.ClientId, oAuthInfo.ClientSecret, oAuthInfo.TenantId));
            ExchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, oAuthInfo.ImpersonateUser);
            ExchangeService.Timeout = 500000;
            //ExchangeService.TraceListener = traceListener;
            //ExchangeService.TraceEnabled = true;
            ExchangeService.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
            Username = oAuthInfo.ImpersonateUser;
        }

        public string GetOAuthToken(string appId, string clientSecret, string tenantId)
		{
            // Using Microsoft.Identity.Client 4.22.0
            var cca = ConfidentialClientApplicationBuilder
                .Create(appId)
                .WithClientSecret(clientSecret)
                .WithTenantId(tenantId)
                .Build();

            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

            //Make the token request
            var authResult = cca.AcquireTokenForClient(ewsScopes).ExecuteAsync().Result;
            return authResult.AccessToken;
        }

        public void ExecuteCall(Action action)
		{
            bool needRetry = true;
            int retryCount = 0;
            do
            {
                try
                {
                    action();
                    needRetry = false;
                }
                catch (ServerBusyException ex)
                {
                    Console.WriteLine($"Server is busy. Retrying after {ex.BackOffMilliseconds/1000}sec");
                    Logger.FileLogger.Warning($"Server is busy. Retrying after {ex.BackOffMilliseconds / 1000}sec");
                    Thread.Sleep(ex.BackOffMilliseconds);
                    needRetry = true; 
                    retryCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception occur while creating mails. Detail: {ex.Message}");
                    Logger.FileLogger.Error($"Exception occur while creating mails. Detail: {ex.Message}");
                    Console.WriteLine($"Retrying after {EWSServiceConstants.RETRY_AFTER / 1000}sec");
                    Logger.FileLogger.Warning($"Retrying after {EWSServiceConstants.RETRY_AFTER / 1000}sec");
                    Thread.Sleep(EWSServiceConstants.RETRY_AFTER);
                    needRetry = true;
                    retryCount++;
                }
            } while (needRetry && retryCount < EWSServiceConstants.RETRY_COUNT);
        }

        public TResult ExecuteCall<TResult>(Func<TResult> function)
        {
            TResult result = default(TResult);
            bool needRetry = true;
            int retryCount = 0;
            do
            {
                try
                {
                    result = function();
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
                    Console.WriteLine($"Exception occur while creating mails. Detail: {ex.Message}");
                    Logger.FileLogger.Error($"Exception occur while creating mails. Detail: {ex.Message}");
                    Console.WriteLine($"Retrying after {EWSServiceConstants.RETRY_AFTER / 1000}sec");
                    Logger.FileLogger.Warning($"Retrying after {EWSServiceConstants.RETRY_AFTER / 1000}sec");
                    Thread.Sleep(EWSServiceConstants.RETRY_AFTER);
                    needRetry = true;
                    retryCount++;
                }
            } while (needRetry && retryCount < EWSServiceConstants.RETRY_COUNT);
            return result;
        }


    }
}
