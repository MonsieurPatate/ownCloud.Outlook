using System;
using Microsoft.Office.Interop.Outlook;

namespace ownCloud.Outlook.CustomEventArgs
{
    public class InspectorCloseEventArgs : EventArgs
    {
        public Inspector Inspector { get; }

        public InspectorCloseEventArgs(Inspector inspector)
        {
            Inspector = inspector;
        }
    }
}