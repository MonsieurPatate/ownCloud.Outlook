using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using AdysTech.CredentialManager;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using RestSharp.Authenticators;
using WebDav;

namespace ownCloud.Outlook
{
    public partial class ThisAddIn
    {
        /// <summary>
        ///     Max size in MB
        /// </summary>
        private int MaxAttachmentSizeMb => MaxAttachmentSize / (1024 * 1024);

        /// <summary>
        ///     Max size in bytes
        /// </summary>
        private const int MaxAttachmentSize = 1024 * 1024 * 10;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var inspectors = Application.Inspectors;
            inspectors.NewInspector += OnCreateNewEmailInspector;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SettingsRibbon();
        }

        private void OnCreateNewEmailInspector(Inspector inspector)
        {
            if (!(inspector.CurrentItem is MailItem mailItem)) return;
            mailItem.BeforeAttachmentAdd += OnBeforeAttachementAdd;
        }

        // ReSharper disable once RedundantAssignment
        private void OnBeforeAttachementAdd(Attachment attachment, ref bool cancel)
        {
            if (attachment.Size <= MaxAttachmentSize) return;

            MessageBox.Show($@"Attachment will be uploaded to fileCloud because the file size exceeds the limit of {MaxAttachmentSizeMb}MB");
            var link = RunTimeContext.Instance.UploadAttachment(attachment);

            var activeInspector = attachment.Application.ActiveInspector();
            var mailItem = (MailItem)activeInspector.CurrentItem;
            mailItem.Body = string.Concat(mailItem.Body, Environment.NewLine, link);

            cancel = true;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
        [SuppressMessage("ReSharper", "RedundantDelegateCreation")]
        [SuppressMessage("ReSharper", "RedundantNameQualifier")]
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}