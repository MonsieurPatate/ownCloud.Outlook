using System;
using Microsoft.Office.Interop.Outlook;

namespace ownCloud.Outlook.InspectorWrappers
{
    public class MailItemWrapper
    {
        private readonly MailItem _mailItem;

        /// <summary>
        ///     Max size in bytes
        /// </summary>
        private const int MaxAttachmentSize = 1024 * 1024 * 10;

        public MailItemWrapper(MailItem mailItem)
        {
            _mailItem = mailItem;
            SubscribeOnEvents();
        }

        private void SubscribeOnEvents()
        {
            _mailItem.BeforeAttachmentAdd += OnBeforeAttachementAdd;
        }

        private void OnBeforeAttachementAdd(Attachment attachment, ref bool cancel)
        {
            // if (attachment.Size <= MaxAttachmentSize) return;
            
            var link = RuntimeContext.Instance.UploadAttachment(attachment);

            var activeInspector = attachment.Application.ActiveInspector();
            var mailItem = (MailItem)activeInspector.CurrentItem;
            mailItem.Body = string.Concat(mailItem.Body, Environment.NewLine, link);

            cancel = true;
        }
    }
}