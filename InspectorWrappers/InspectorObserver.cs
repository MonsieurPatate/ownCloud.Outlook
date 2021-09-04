using System;
using System.Collections.Generic;
using ownCloud.Outlook.CustomEventArgs;
using Microsoft.Office.Interop.Outlook;

namespace ownCloud.Outlook.InspectorWrappers
{
    public class InspectorObserver
    {
        // Connect class-level Instance Variables
        // Outlook inspectors collection
        private Inspectors _inspectors;

        // Collection of tracked inspector windows              
        private readonly List<InspectorWrapper> _inspectorWindows = new List<InspectorWrapper>();

        public InspectorObserver(Inspectors inspectors)
        {
            _inspectors = inspectors;
            _inspectors.NewInspector += inspectors_NewInspector;
        }

        // NewInspector event creates new instance of InspectorWrapper
        public void inspectors_NewInspector(Inspector inspector)
        {
            // Check to see if this is a new window you don't
            // already track
            var existingWindow = FindOutlookInspector(inspector);
            if (existingWindow == null)
            {
                AddInspector(inspector);
            }
        }

        // Adds an instance of **InspectorWrapper** class
        private void AddInspector(Inspector inspector)
        {
            var window = new InspectorWrapper(inspector);
            window.Close += OnInspectorWindowClose;
            _inspectorWindows.Add(window);
        }

        private void OnInspectorWindowClose(object sender, InspectorCloseEventArgs e)
        {
            var inspector = FindOutlookInspector(e.Inspector);
            if (inspector == null)
            {
                return;
            }

            _inspectorWindows.Remove(inspector);
        }

        // Looks up the window wrapper for a given Inspector 
        // window object
        private InspectorWrapper FindOutlookInspector(Inspector window)
        {
            foreach (var inspector in _inspectorWindows)
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