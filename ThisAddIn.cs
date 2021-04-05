using System;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Office.Interop.Outlook;
using WebDav;

namespace ownCloud.Outlook
{
    public partial class ThisAddIn
    {
        private static IWebDavClient _webDavClient = new WebDavClient();

        /// <summary>
        ///     Max size in MB
        /// </summary>
        private const int MaxAttachmentSize = 1024 * 1024 * 10;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var inspectors = Application.Inspectors;
            inspectors.NewInspector += OnCreateNewEmailInspector;
        }

        private void OnCreateNewEmailInspector(Inspector inspector)
        {
            if (!(inspector.CurrentItem is MailItem mailItem)) return;
            mailItem.BeforeAttachmentAdd += OnBeforeAttachementAdd;
        }

        private void OnBeforeAttachementAdd(Attachment attachment, ref bool cancel)
        {
            if (attachment.Size <= MaxAttachmentSize) return;
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
