using System;
using System.Windows.Forms.VisualStyles;
using Outlook_ = Microsoft.Office.Interop.Outlook;

namespace ownCloud.Outlook.InspectorWrapper
{
    public class OutlookInspector
    {
        public Outlook_.Inspector Window => _mWindow;

        // OutlookInspector class-level instance variables 
        // wrapped window object
        private Outlook_.Inspector _mWindow;

        // Use these instance variables to handle item-level events
        // wrapped MailItem
        private Outlook_.MailItem _mMail;

        // wrapped AppointmentItem        
        private Outlook_.AppointmentItem _mAppointment;

        // wrapped ContactItem
        private Outlook_.ContactItem _mContact;

        // wrapped TaskItem      
        private Outlook_.TaskItem _mTask;

        public EventHandler Close;

        // OutlookInspector constructor
        public OutlookInspector(Outlook_.Inspector inspector)
        {
            _mWindow = inspector;

            // Hook up the close event
            ((Outlook_.InspectorEvents_Event) inspector).Close +=
                OutlookInspectorWindow_Close;

            // Hook up item-level events as needed
            OutlookItem olItem = new OutlookItem(inspector.CurrentItem);
            if (olItem.Class == Outlook_.OlObjectClass.olContact)
            {
                /*m_Contact = olItem.InnerObject as Outlook_.ContactItem;
                m_Contact.Open +=
                    m_Contact_Open;
                m_Contact.PropertyChange +=
                    m_Contact_PropertyChange;
                m_Contact.CustomPropertyChange +=
                    m_Contact_CustomPropertyChange;*/
            }
        }

        // Event Handler for the inspector close event.
        private void OutlookInspectorWindow_Close()
        {
            // Unhook events from any item-level instance variables
            /*m_Contact.Open -=
                Outlook_.ItemEvents_10_OpenEventHandler(
                    m_Contact_Open);
            m_Contact.PropertyChange -=
                Outlook_.ItemEvents_10_PropertyChangeEventHandler(
                    m_Contact_PropertyChange);
            m_Contact.CustomPropertyChange -=
                Outlook_.ItemEvents_10_CustomPropertyChangeEventHandler(
                    m_Contact_CustomPropertyChange);
            ((Outlook_.ItemEvents_Event) m_Contact).Close -=
                Outlook_.ItemEvents_CloseEventHandler(
                    m_Contact_Close);*/

            // Unhook events from the window
            ((Outlook_.InspectorEvents_Event) _mWindow).Close -= OutlookInspectorWindow_Close;

            // Raise the OutlookInspector close event
            Close?.Invoke(this, EventArgs.Empty);
                
            // Release item-level instance variables
            _mMail = null;
            _mAppointment = null;
            _mContact = null;
            _mTask = null;
            _mWindow = null;
        }
    }
}