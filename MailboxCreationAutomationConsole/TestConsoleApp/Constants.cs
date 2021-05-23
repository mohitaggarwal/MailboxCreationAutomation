using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsoleApp
{
	public static class Constants
	{

		// First full backup
		public const string FirstFullBackupRootFolder = "First Full Backup";
		public const string FirstFullBackupToRecipient = "LidiaH@dpo365backup.onmicrosoft.com";
		public const string FirstFullBackupFromRecipient = "Haripriya@hpdatapro.onmicrosoft.com";
		public const string FirstFullBackupMailSubject = "First full backup mail";
		public const string FirstFullBackupEventSubject = "First full backup event";
		public const string FirstFullBackupEventBody = "First full backup event";
		public const string FirstFullBackupContactName= "First full backup";
		public const string FirstFullBackupContactEmail = "firstfullbackup@gmail.com";

		// Incr backup
		public const string IncrBackupRootFolder = "Incr Backup_{0}";
		public const string IncrBackupToRecipient = "LidiaH@dpo365backup.onmicrosoft.com";
		public const string IncrBackupFromRecipient = "Haripriya@hpdatapro.onmicrosoft.com";
		public const string IncrBackupMailSubject = "Incr Backup_{0} mail";
		public const string IncrBackupEventSubject = "Incr Backup_{0} event";
		public const string IncrBackupEventBody = "Incr Backup_{0} event";
		public const string IncrBackupContactName = "Incr Backup_{0}";
		public const string IncrBackupContactEmail = "incrbackup{0}@gmail.com";

	}
}
