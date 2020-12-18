using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxCreationAutomation
{
	public class MailboxMails
	{
		EWSServiceWrapper _EWSServiceWrapper;

		public MailboxMails(EWSServiceWrapper eWSServiceWrapper)
		{
			_EWSServiceWrapper = eWSServiceWrapper;
		}

		public void CreateAttachment(string fileName, int sizeInMB)
		{
			using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
			{
				fs.SetLength(sizeInMB * 1024 * 1024);
			}
		}

		private void MailAttachments(string directory, int attachments, int sizeInMB, ref EmailMessage message)
		{
			for(int i = 1; i <= attachments; i++)
			{
				string attachmentName = $"Attachment_{i}_sizeInMB_{sizeInMB}.txt";
				string fileName = Path.Combine(directory, attachmentName);
				CreateAttachment(fileName, sizeInMB);
				// Add a file attachment by using a stream, and specify the name of the attachment.
				// The email attachment is named FourthAttachment.txt.
				FileStream theStream = new FileStream(fileName, FileMode.OpenOrCreate);
				// In this example, theStream is a Stream object that represents the content of the file to attach.
				message.Attachments.AddFileAttachment(attachmentName, theStream);
			}
		}

		public void CreateMails(MailsToCreate mailsToCreate, string folderId)
		{
			for(int i = 1; i <= mailsToCreate.Mails; i++)
			{

				bool needRetry = true;
				do
				{

					string directory = Path.Combine(System.Environment.
													 GetFolderPath(
														 Environment.SpecialFolder.CommonApplicationData),
													"TempAttachmentFolder");
					DirectoryInfo di = Directory.CreateDirectory(directory);
					try
					{

						EmailMessage message = new EmailMessage(_EWSServiceWrapper.ExchangeService);
						// Set properties on the email message.
						message.ToRecipients.Add(_EWSServiceWrapper.Username);
						message.From = new EmailAddress { Address = mailsToCreate.From };
						message.Subject = $"Mail message {i} with attachments {mailsToCreate.Attachments} and attachment size {mailsToCreate.AttachmentSizeInMB}";
						MailAttachments(directory, mailsToCreate.Attachments, mailsToCreate.AttachmentSizeInMB, ref message);
						//message.Body = "(1) Buy pizza, (2) Eat pizza";
						ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
						message.SetExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);
						// This method call results in a CreateItem call to EWS.
						message.Save(folderId);
						needRetry = false;
					}
					catch (ServerBusyException ex)
					{
						Console.WriteLine($"Server is busy. Retrying after {ex.BackOffMilliseconds / 1000}sec");
						Thread.Sleep(ex.BackOffMilliseconds);
						needRetry = true;
					}
					finally
					{
						di.Delete(true);
					}
				} while (needRetry);
			}
		}
	}
}
