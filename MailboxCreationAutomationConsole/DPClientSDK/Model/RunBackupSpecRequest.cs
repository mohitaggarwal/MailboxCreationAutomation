using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DPClientSDK.Model
{
	public class RunBackupSpecRequest
	{
		[JsonProperty(PropertyName = "specificationName")]
		public string SpecificationName { get; set; }

		[JsonProperty(PropertyName = "applicationType")]
		public string ApplicationType { get; set; }

		[JsonProperty(PropertyName = "mode")]
		public string Mode { get; set; }

		[JsonProperty(PropertyName = "load")]
		public string Load { get; set; }

		[JsonProperty(PropertyName = "monitor")]
		public string Monitor { get; set; }
	}
}
