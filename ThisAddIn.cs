using System;
using System.Diagnostics.CodeAnalysis;
using System.Net;
using ownCloud.Outlook.InspectorWrappers;

namespace ownCloud.Outlook
{
    public partial class ThisAddIn
    {
        private InspectorObserver _inspectorObserver;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // enable establish connection to WebDav via https
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            _inspectorObserver = new InspectorObserver(Application.Inspectors);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SettingsRibbon();
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