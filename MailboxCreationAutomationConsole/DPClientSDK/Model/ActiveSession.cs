using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DPClientSDK.Model
{
	public class ActiveSession
	{
		[JsonProperty(PropertyName = "sess_name")]
		public string SessionName { get; set; }

		[JsonProperty(PropertyName = "sess_datalist")]
		public string BackupSpec { get; set; }

		[JsonProperty(PropertyName = "application_type")]
		public string ApplicationType { get; set; }

		[JsonProperty(PropertyName = "sess_start_time")]
		public long SessionStartTime { get; set; }

		[JsonProperty(PropertyName = "sess_type")]
		public int SessionType { get; set; }

		[JsonProperty(PropertyName = "sess_status")]
		public int Status { get; set; }

		[JsonProperty(PropertyName = "backup_type")]
		public int BackupType { get; set; }

		[JsonProperty(PropertyName = "size")]
		public int Size { get; set; }

		[JsonProperty(PropertyName = "duration")]
		public int Duration { get; set; }
	}

	public class ActiveSessions
	{
		[JsonProperty(PropertyName = "sessions")]
		public List<ActiveSession> Sessions { get; set; }
	}
}
