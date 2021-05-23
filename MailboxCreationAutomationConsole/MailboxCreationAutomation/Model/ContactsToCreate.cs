using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation.Model
{
	public class ContactsToCreate
	{
		public ContactToCreate ContactToCreate { get; set; }
		public int Count { get; set; }
	}

	public class ContactToCreate
	{
		public string GivenName { get; set; }
		public string MiddleName { get; set; }
		public string Surname { get; set; }
		public string CompanyName { get; set; }
		public string BussinessPhone { get; set; }
		public string HomePhone { get; set; }
		public string CarPhone { get; set; }
		public string EmailAddress { get; set; }
		public AddressToCreate HomeAddress { get; set; }
		public AddressToCreate BussinessAddress { get; set; }
		public string PhotoPath { get; set; }
	}

	public class AddressToCreate
	{
		public string Street { get; set; }
		public string City { get; set; }
		public string State { get; set; }
		public string PostalCode { get; set; }
		public string CountryOrRegion { get; set; }
	}
}
