using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DPClientSDK.Model
{
	public class RunBackupSpecResponse
	{
		[JsonProperty(PropertyName = "status")]
		public string Status { get; set; }

		[JsonProperty(PropertyName = "session_id")]
		public string SessionId { get; set; }
	}
}
