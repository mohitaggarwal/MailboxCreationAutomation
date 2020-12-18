using Microsoft.Exchange.WebServices.Data;
using System;
using System.Net;
using System.Threading;

namespace MailboxCreationAutomation
{
    public class EWSServiceWrapper
    {
        public ExchangeService ExchangeService { get; set; }
        public string Username { get; set; }

        public EWSServiceWrapper(string username, string password)
		{
            TraceListener traceListener = new TraceListener();
            ExchangeService = new ExchangeService(ExchangeVersion.Exchange2013);
            ExchangeService.Credentials = new WebCredentials(username, password);
            ExchangeService.Url = new Uri(EWSServiceConstants.EWS_URL);
            ExchangeService.Timeout = 500000;
            ExchangeService.TraceListener = traceListener;
            ExchangeService.TraceEnabled = true;
            ExchangeService.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
            Username = username;
        }

        public void ExecuteCall(Action action)
		{
            bool needRetry = true;
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
                    Thread.Sleep(ex.BackOffMilliseconds);
                    needRetry = true;
                }
            } while (needRetry);
		}


    }
}
