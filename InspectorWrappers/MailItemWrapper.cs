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
            var bodyIndex = mailItem.HTMLBody.IndexOf("</body>", StringComparison.InvariantCultureIgnoreCase);
            mailItem.HTMLBody = mailItem.HTMLBody.Insert(bodyIndex, $"<br><a href=\"{link}\"/>{link}</br>");
            // mailItem.Body.Insert(bodyLength, string.Concat(Environment.NewLine, link));
            cancel = true;
        }
    }
}