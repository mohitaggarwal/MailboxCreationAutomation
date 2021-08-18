using MailboxCreationAutomation.Model;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
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

		private MessageBody GetMailBody(string path)
		{
			MessageBody msgBody = null;
			string ext = Path.GetExtension(path);
			if (File.Exists(path) && (ext.Equals(".html", StringComparison.OrdinalIgnoreCase)
				|| ext.Equals(".htm", StringComparison.OrdinalIgnoreCase)))
			{
				msgBody = new MessageBody();
				msgBody.BodyType = BodyType.HTML;
				msgBody.Text = File.ReadAllText(path);
			}
			return msgBody;
		}

		private string GetAttachmentTemplate(List<AttachmentsToCreate> attachmentsToCreates)
		{
			StringBuilder attachmentsTemplate = new StringBuilder();
			if (attachmentsToCreates != null && attachmentsToCreates.Count > 0)
			{
				foreach (var attachmentToCreate in attachmentsToCreates)
				{
					attachmentsTemplate.AppendLine($"'{attachmentToCreate.Count}' attachments of size '{attachmentToCreate.AttachmentSizeInKB}Kb', ");
				}
			}
			else
			{
				attachmentsTemplate.AppendLine($"zero attachments.");
			}
			return attachmentsTemplate.ToString(); ;
		}

		private void CreateAttachment(string fileName, int sizeInKB)
		{
			using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
			{
				fs.SetLength(sizeInKB *  1024);
			}
		}

		private void MailAttachments(string directory, AttachmentsToCreate attachmentsToCreate, string prefix, ref EmailMessage message)
		{
			if (attachmentsToCreate != null)
			{
				for (int i = 1; i <= attachmentsToCreate.Count; i++)
				{
					string attachmentName = $"{prefix}_{i}_sizeInKB_{attachmentsToCreate.AttachmentSizeInKB}.txt";
					string fileName = Path.Combine(directory, attachmentName);
					CreateAttachment(fileName, attachmentsToCreate.AttachmentSizeInKB);
					// Add a file attachment by using a stream, and specify the name of the attachment.
					// The email attachment is named FourthAttachment.txt.
					FileStream theStream = new FileStream(fileName, FileMode.OpenOrCreate);
					// In this example, theStream is a Stream object that represents the content of the file to attach.
					message.Attachments.AddFileAttachment(attachmentName, theStream);
				}
			}
		}

		private void CreateMails(MailsToCreate mailsToCreate, string folderId, string prefix, int number)
		{
			string directory = $"{_EWSServiceWrapper.Username}_{DateTime.Now.ToString("yyyy - MM - dd HH - mm - ss - fff")}";
			DirectoryInfo di = Directory.CreateDirectory(directory);
			try
			{
				EmailMessage message = new EmailMessage(_EWSServiceWrapper.ExchangeService);
				// Set properties on the email message.
				if (mailsToCreate.To != null)
				{
					foreach (var toRecip in mailsToCreate.To)
					{
						message.ToRecipients.Add(toRecip);
					}
				}
				else
				{
					message.ToRecipients.Add(_EWSServiceWrapper.Username);
				}
				message.From = new EmailAddress { Address = mailsToCreate.From };
				message.Subject = $"{mailsToCreate.Subject}_{number}_{prefix} with attachments {mailsToCreate.AttachmentsToCreateList?.Sum(x => x.Count)} and attachment size {mailsToCreate.AttachmentsToCreateList?.Sum(x => x.Count * x.AttachmentSizeInKB)}KB";
				MessageBody body = GetMailBody(mailsToCreate.BodyPath);
				if (body != null)
				{
					message.Body = body;
				}
				else
				{
					message.Body = string.Format(EWSServiceConstants.MAIL_BODY_TEMPLATE, message.Subject, GetAttachmentTemplate(mailsToCreate.AttachmentsToCreateList));
				}
				if (mailsToCreate.AttachmentsToCreateList != null)
				{
					foreach (var attachmentToCreate in mailsToCreate.AttachmentsToCreateList)
					{
						MailAttachments(directory, attachmentToCreate, $"{prefix}_Attach_Set_{number}", ref message);
					}
				}
				ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
				message.SetExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);
				message.Save(folderId);
				Logger.FileLogger.Info($"Message with subject '{message.Subject}' and id '{message.Id.UniqueId}' created successfully.");
			}
			finally
			{
				if (di.Exists)
				{
					di.Delete(true);
				}
			}
		}

		public void CreateMails(MailsToCreate mailsToCreate, string folderId, string prefix)
		{
			if (mailsToCreate != null)
			{
				for (int i = 1; i <= mailsToCreate.Count; i++)
				{
					_EWSServiceWrapper.ExecuteCall(() => CreateMails(mailsToCreate, folderId, prefix, i));
				}
			}
		}

		public List<Item> GetItems(string folderId)
		{
			List<Item> items = new List<Item>();
			PropertySet propertySet = new PropertySet
			{
				ItemSchema.Id,
				ItemSchema.Subject,
				ItemSchema.Size
			};
			bool isMoreAvailable = false;
			string syncState = string.Empty;
			do
			{
				var itemChanges = _EWSServiceWrapper.ExecuteCall(() => 
									_EWSServiceWrapper.ExchangeService.SyncFolderItems(
										folderId, propertySet, null, 512, SyncFolderItemsScope.NormalItems, syncState));
				syncState = itemChanges.SyncState;
				items.AddRange(itemChanges.Select(x => x.Item));
				isMoreAvailable = itemChanges.MoreChangesAvailable;
			} while (isMoreAvailable);

			return items;
		}

		public ExportItemsResponse ExportItem(ItemId itemId)
		{
			return _EWSServiceWrapper.ExecuteCall(() =>
						_EWSServiceWrapper.ExchangeService
						.ExportItems(new List<ItemId> { itemId}).FirstOrDefault());
		}

		public List<ExportItemsResponse> ExportItem(List<ItemId> itemIds)
		{
			return _EWSServiceWrapper.ExecuteCall(() =>
						_EWSServiceWrapper.ExchangeService
						.ExportItems(itemIds).ToList());
		}

		public void DeleteMail(string mailId)
		{
			EmailMessage emailMessage = EmailMessage.Bind(_EWSServiceWrapper.ExchangeService, mailId);
			emailMessage.Delete(DeleteMode.HardDelete);
		}
	}
}
