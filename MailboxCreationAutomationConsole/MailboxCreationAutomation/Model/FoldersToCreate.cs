﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class FoldersToCreate
	{
		public string Prefix { get; set; }
		public int Count { get; set; }
		public int Levels { get; set; }
		public List<MailsToCreate> MailsToCreateList { get; set; }
	}
}
