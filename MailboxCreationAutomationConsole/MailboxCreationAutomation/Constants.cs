using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public static class EWSServiceConstants
	{
		public static string EWS_URL = "https://outlook.office365.com/EWS/Exchange.asmx";
		public static string MAIL_SUBJECT = "Mail message ";
		public static int RETRY_AFTER = 30000; // 30 sec
		public static int RETRY_COUNT = 20;
		public static string MAIL_BODY_TEMPLATE = @"Creating a mail with
											Subject: '{0}'
											and
											Attachments: {1}";

		public static string ATTACHENT_TEMPLATE = "'{0}' attachments of size '{1}Kb'";
	}
}
