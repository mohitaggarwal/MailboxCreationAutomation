using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class Contacts
	{
		EWSServiceWrapper _EWSServiceWrapper;

		public Contacts(EWSServiceWrapper eWSServiceWrapper)
		{
			_EWSServiceWrapper = eWSServiceWrapper;
		}

		private void CreateContacts(ContactsToCreate contactsToCreate, string folderId, string prefix, int number)
		{
			if (contactsToCreate != null && contactsToCreate.ContactToCreate != null)
			{
				Contact contact = new Contact(_EWSServiceWrapper.ExchangeService);
				contact.GivenName = contactsToCreate.ContactToCreate.GivenName;
				contact.MiddleName = contactsToCreate.ContactToCreate.MiddleName;
				contact.Surname = contactsToCreate.ContactToCreate.Surname;
				contact.DisplayName = $"{contact.GivenName}_{number}_{prefix}";
				contact.FileAsMapping = FileAsMapping.SurnameCommaGivenName;
				contact.CompanyName = contactsToCreate.ContactToCreate.CompanyName;
				contact.PhoneNumbers[PhoneNumberKey.BusinessPhone] = contactsToCreate.ContactToCreate.BussinessPhone;
				contact.PhoneNumbers[PhoneNumberKey.HomePhone] = contactsToCreate.ContactToCreate.HomePhone;
				contact.PhoneNumbers[PhoneNumberKey.CarPhone] = contactsToCreate.ContactToCreate.CarPhone;
				contact.EmailAddresses[EmailAddressKey.EmailAddress1] = new EmailAddress(contactsToCreate.ContactToCreate.EmailAddress);

				if (contactsToCreate.ContactToCreate.HomeAddress != null)
				{
					// Specify the home address.
					PhysicalAddressEntry physicalAddress = new PhysicalAddressEntry();
					physicalAddress.Street = contactsToCreate.ContactToCreate.HomeAddress.Street;
					physicalAddress.City = contactsToCreate.ContactToCreate.HomeAddress.City;
					physicalAddress.State = contactsToCreate.ContactToCreate.HomeAddress.State;
					physicalAddress.PostalCode = contactsToCreate.ContactToCreate.HomeAddress.PostalCode;
					physicalAddress.CountryOrRegion = contactsToCreate.ContactToCreate.HomeAddress.CountryOrRegion;
					contact.PhysicalAddresses[PhysicalAddressKey.Home] = physicalAddress;
				}

				if (contactsToCreate.ContactToCreate.BussinessAddress != null)
				{
					// Specify the business address.
					PhysicalAddressEntry bussinessAddress = new PhysicalAddressEntry();
					bussinessAddress.Street = contactsToCreate.ContactToCreate.BussinessAddress.Street;
					bussinessAddress.City = contactsToCreate.ContactToCreate.BussinessAddress.City;
					bussinessAddress.State = contactsToCreate.ContactToCreate.BussinessAddress.State;
					bussinessAddress.PostalCode = contactsToCreate.ContactToCreate.BussinessAddress.PostalCode;
					bussinessAddress.CountryOrRegion = contactsToCreate.ContactToCreate.BussinessAddress.CountryOrRegion;
					contact.PhysicalAddresses[PhysicalAddressKey.Business] = bussinessAddress;
				}

				if (!string.IsNullOrEmpty(contactsToCreate.ContactToCreate.PhotoPath)
					&& File.Exists(contactsToCreate.ContactToCreate.PhotoPath))
				{
					FileAttachment atattach = contact.Attachments.AddFileAttachment(contactsToCreate.ContactToCreate.PhotoPath);
					atattach.IsContactPhoto = true;
				}

				contact.Save(folderId);
			}
		}

		public void CreateContacts(ContactsToCreate contactsToCreate, string folderId, string prefix)
		{
			if (contactsToCreate != null)
			{
				for (int i = 1; i <= contactsToCreate.Count; i++)
				{
					_EWSServiceWrapper.ExecuteCall(() => CreateContacts(contactsToCreate, folderId, prefix, i));
				}
			}
		}

		public void DeleteEvent(string contactId)
		{
			Contact contact = Contact.Bind(_EWSServiceWrapper.ExchangeService, contactId);
			contact.Delete(DeleteMode.HardDelete);
		}

	}
}
