using System;
using System.Collections.Generic;
using Outlook_ = Microsoft.Office.Interop.Outlook;

namespace ownCloud.Outlook.InspectorWrapper
{
    public class InspectorObserver
    {
        // Connect class-level Instance Variables
        // Outlook inspectors collection
        private Outlook_.Inspectors inspectors;

        // Collection of tracked inspector windows              
        private readonly List<OutlookInspector> _inspectorWindows = new List<OutlookInspector>();

        // NewInspector event creates new instance of OutlookInspector
        public void inspectors_NewInspector(Outlook_.Inspector inspector)
        {
            // Check to see if this is a new window you don't
            // already track
            OutlookInspector existingWindow = FindOutlookInspector(inspector);
            if (existingWindow == null)
            {
                AddInspector(inspector);
            }
        }

        // Adds an instance of **OutlookInspector** class
        private void AddInspector(Outlook_.Inspector inspector)
        {
            if ((inspector.CurrentItem is Outlook_.MailItem mailItem))
            {
                return;
            }

            OutlookInspector window = new OutlookInspector(inspector);
            _inspectorWindows.Add(window);
            // window.Close += WrappedInspectorWindow_Close;
        }

        // Looks up the window wrapper for a given Inspector 
        // window object
        private OutlookInspector FindOutlookInspector(object window)
        {
            foreach (OutlookInspector inspector in _inspectorWindows)
            {
                if (inspector.Window == window)
                {
                    return inspector;
                }
            }
            return null;
        }
    }
}