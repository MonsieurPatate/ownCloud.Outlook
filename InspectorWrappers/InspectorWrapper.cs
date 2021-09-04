using System;
using ownCloud.Outlook.CustomEventArgs;
using Microsoft.Office.Interop.Outlook;

namespace ownCloud.Outlook.InspectorWrappers
{
    public class InspectorWrapper
    {
        public Inspector Window => _inspector;

        // InspectorWrapper class-level instance variables 
        // wrapped window object
        private Inspector _inspector;

        private MailItemWrapper _mailItem;

        public EventHandler<InspectorCloseEventArgs> Close;


        // InspectorWrapper constructor
        public InspectorWrapper(Inspector inspector)
        {
            _inspector = inspector;

            // Hook up the close event
            ((InspectorEvents_Event)inspector).Close += OnInspectorWindowClose;

            if (inspector.CurrentItem is MailItem mailItem)
            {
                _mailItem = new MailItemWrapper(mailItem);
            }
        }

        // Event Handler for the inspector close event.
        private void OnInspectorWindowClose()
        {
            // Unhook events from the window
            ((InspectorEvents_Event) _inspector).Close -= OnInspectorWindowClose;

            // Raise the InspectorWrapper close event
            Close?.Invoke(this, new InspectorCloseEventArgs(_inspector));
                
            // Release item-level instance variables
            _mailItem = null;
            _inspector = null;
        }
    }
}