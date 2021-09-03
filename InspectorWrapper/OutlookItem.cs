using System;
using System.Diagnostics;
using System.Reflection;
using Outlook_ = Microsoft.Office.Interop.Outlook;

namespace ownCloud.Outlook.InspectorWrapper
{
    internal class OutlookItem
    {
        private readonly Type m_type; // type for the Outlook_ item 
        private readonly object[] m_args; // dummy argument array
        private Type m_typeOlObjectClass;

        #region OutlookItem Constants

        private const string OlActions = "Actions";
        private const string OlApplication = "Application";
        private const string OlAttachments = "Attachments";
        private const string OlBillingInformation = "BillingInformation";
        private const string OlBody = "Body";
        private const string OlCategories = "Categories";
        private const string OlClass = "Class";
        private const string OlClose = "Close";
        private const string OlCompanies = "Companies";
        private const string OlConversationIndex = "ConversationIndex";
        private const string OlConversationTopic = "ConversationTopic";
        private const string OlCopy = "Copy";
        private const string OlCreationTime = "CreationTime";
        private const string OlDisplay = "Display";
        private const string OlDownloadState = "DownloadState";
        private const string OlEntryID = "EntryID";
        private const string OlFormDescription = "FormDescription";
        private const string OlGetInspector = "GetInspector";
        private const string OlImportance = "Importance";
        private const string OlIsConflict = "IsConflict";
        private const string OlItemProperties = "ItemProperties";
        private const string OlLastModificationTime = "LastModificationTime";
        private const string OlLinks = "Links";
        private const string OlMarkForDownload = "MarkForDownload";
        private const string OlMessageClass = "MessageClass";
        private const string OlMileage = "Mileage";
        private const string OlMove = "Move";
        private const string OlNoAging = "NoAging";
        private const string OlOutlookInternalVersion = "OutlookInternalVersion";
        private const string OlOutlookVersion = "OutlookVersion";
        private const string OlParent = "Parent";
        private const string OlPrintOut = "PrintOut";
        private const string OlPropertyAccessor = "PropertyAccessor";
        private const string OlSave = "Save";
        private const string OlSaveAs = "SaveAs";
        private const string OlSaved = "Saved";
        private const string OlSensitivity = "Sensitivity";
        private const string OlSession = "Session";
        private const string OlShowCategoriesDialog = "ShowCategoriesDialog";
        private const string OlSize = "Size";
        private const string OlSubject = "Subject";
        private const string OlUnRead = "UnRead";
        private const string OlUserProperties = "UserProperties";

        #endregion

        #region Constructor

        public OutlookItem(object item)
        {
            InnerObject = item;
            m_type = InnerObject.GetType();
            m_args = new object[] { };
        }

        #endregion

        #region Public Methods and Properties

        public Outlook_.Actions Actions => GetPropertyValue(OlActions) as Outlook_.Actions;

        public Outlook_.Application Application => GetPropertyValue(OlApplication) as Outlook_.Application;

        public Outlook_.Attachments Attachments => GetPropertyValue(OlAttachments) as Outlook_.Attachments;

        public string BillingInformation
        {
            get => GetPropertyValue(OlBillingInformation).ToString();
            set => SetPropertyValue(OlBillingInformation, value);
        }

        public string Body
        {
            get => GetPropertyValue(OlBody).ToString();
            set => SetPropertyValue(OlBody, value);
        }

        public string Categories
        {
            get => GetPropertyValue(OlCategories).ToString();
            set => SetPropertyValue(OlCategories, value);
        }


        public void Close(Outlook_.OlInspectorClose SaveMode)
        {
            object[] MyArgs = {SaveMode};
            CallMethod(OlClose, MyArgs);
        }

        public string Companies
        {
            get => GetPropertyValue(OlCompanies).ToString();
            set => SetPropertyValue(OlCompanies, value);
        }

        public Outlook_.OlObjectClass Class
        {
            get
            {
                if (m_typeOlObjectClass == null)
                {
                    // Note: instantiate dummy ObjectClass enumeration to get type.
                    //       type = System.Type.GetType("Outlook_.OlObjectClass") doesn't seem to work
                    Outlook_.OlObjectClass objClass = Outlook_.OlObjectClass.olAction;
                    m_typeOlObjectClass = objClass.GetType();
                }

                return (Outlook_.OlObjectClass) Enum.ToObject(m_typeOlObjectClass, GetPropertyValue(OlClass));
            }
        }

        public string ConversationIndex => GetPropertyValue(OlConversationIndex).ToString();

        public string ConversationTopic => GetPropertyValue(OlConversationTopic).ToString();

        public object Copy()
        {
            return CallMethod(OlCopy);
        }

        public DateTime CreationTime => (DateTime) GetPropertyValue(OlCreationTime);

        public void Display()
        {
            CallMethod(OlDisplay);
        }

        public Outlook_.OlDownloadState DownloadState => (Outlook_.OlDownloadState) GetPropertyValue(OlDownloadState);

        public string EntryID => GetPropertyValue(OlEntryID).ToString();

        public Outlook_.FormDescription FormDescription => (Outlook_.FormDescription) GetPropertyValue(OlFormDescription);


        public object InnerObject { get; }

        public Outlook_.Inspector GetInspector => GetPropertyValue(OlGetInspector) as Outlook_.Inspector;

        public Outlook_.OlImportance Importance
        {
            get => (Outlook_.OlImportance) GetPropertyValue(OlImportance);
            set => SetPropertyValue(OlImportance, value);
        }

        public bool IsConflict => (bool) GetPropertyValue(OlIsConflict);

        public Outlook_.ItemProperties ItemProperties => (Outlook_.ItemProperties) GetPropertyValue(OlItemProperties);

        public DateTime LastModificationTime => (DateTime) GetPropertyValue(OlLastModificationTime);

        public Outlook_.Links Links => GetPropertyValue(OlLinks) as Outlook_.Links;

        public Outlook_.OlRemoteStatus MarkForDownload
        {
            get => (Outlook_.OlRemoteStatus) GetPropertyValue(OlMarkForDownload);
            set => SetPropertyValue(OlMarkForDownload, value);
        }

        public string MessageClass
        {
            get => GetPropertyValue(OlMessageClass).ToString();
            set => SetPropertyValue(OlMessageClass, value);
        }

        public string Mileage
        {
            get => GetPropertyValue(OlMileage).ToString();
            set => SetPropertyValue(OlMileage, value);
        }

        public object Move(Outlook_.Folder DestinationFolder)
        {
            object[] myArgs = {DestinationFolder};
            return CallMethod(OlMove, myArgs);
        }

        public bool NoAging
        {
            get => (bool) GetPropertyValue(OlNoAging);
            set => SetPropertyValue(OlNoAging, value);
        }

        public long OutlookInternalVersion => (long) GetPropertyValue(OlOutlookInternalVersion);

        public string OutlookVersion => GetPropertyValue(OlOutlookVersion).ToString();

        public Outlook_.Folder Parent => GetPropertyValue(OlParent) as Outlook_.Folder;

        public Outlook_.PropertyAccessor PropertyAccessor => GetPropertyValue(OlPropertyAccessor) as Outlook_.PropertyAccessor;

        public void PrintOut()
        {
            CallMethod(OlPrintOut);
        }

        public void Save()
        {
            CallMethod(OlSave);
        }

        public void SaveAs(string path, Outlook_.OlSaveAsType type)
        {
            object[] myArgs = {path, type};
            CallMethod(OlSaveAs, myArgs);
        }

        public bool Saved => (bool) GetPropertyValue(OlSaved);

        public Outlook_.OlSensitivity Sensitivity
        {
            get => (Outlook_.OlSensitivity) GetPropertyValue(OlSensitivity);
            set => SetPropertyValue(OlSensitivity, value);
        }

        public Outlook_.NameSpace Session => GetPropertyValue(OlSession) as Outlook_.NameSpace;

        public void ShowCategoriesDialog()
        {
            CallMethod(OlShowCategoriesDialog);
        }

        public long Size => (long) GetPropertyValue(OlSize);

        public string Subject
        {
            get => GetPropertyValue(OlSubject).ToString();
            set => SetPropertyValue(OlSubject, value);
        }

        public bool UnRead
        {
            get => (bool) GetPropertyValue(OlUnRead);
            set => SetPropertyValue(OlUnRead, value);
        }

        public Outlook_.UserProperties UserProperties => GetPropertyValue(OlUserProperties) as Outlook_.UserProperties;

        #endregion

        #region Private Helper Functions

        private object GetPropertyValue(string propertyName)
        {
            try
            {
                // An invalid property name exception is propagated to client
                return m_type.InvokeMember(
                    propertyName,
                    BindingFlags.Public | BindingFlags.GetField | BindingFlags.GetProperty,
                    null,
                    InnerObject,
                    m_args);
            }
            catch (SystemException ex)
            {
                Debug.WriteLine("OutlookItem: GetPropertyValue for {0} Exception: {1} ", propertyName, ex.Message);
                throw;
            }
        }

        private void SetPropertyValue(string propertyName, object propertyValue)
        {
            try
            {
                m_type.InvokeMember(
                    propertyName,
                    BindingFlags.Public | BindingFlags.SetField | BindingFlags.SetProperty,
                    null,
                    InnerObject,
                    new[] {propertyValue});
            }
            catch (SystemException ex)
            {
                Debug.WriteLine("OutlookItem: SetPropertyValue for {0} Exception: {1} ", propertyName, ex.Message);
                throw;
            }
        }

        private object CallMethod(string methodName)
        {
            try
            {
                // An invalid property name exception is propagated to client
                return m_type.InvokeMember(
                    methodName,
                    BindingFlags.Public | BindingFlags.InvokeMethod,
                    null,
                    InnerObject,
                    m_args);
            }
            catch (SystemException ex)
            {
                Debug.WriteLine("OutlookItem: CallMethod for {0} Exception: {1} ", methodName, ex.Message);
                throw;
            }
        }

        private object CallMethod(string methodName, object[] args)
        {
            try
            {
                // An invalid property name exception is propagated to client
                return m_type.InvokeMember(
                    methodName,
                    BindingFlags.Public | BindingFlags.InvokeMethod,
                    null,
                    InnerObject,
                    args);
            }
            catch (SystemException ex)
            {
                Debug.WriteLine("OutlookItem: CallMethod for {0} Exception: {1} ", methodName, ex.Message);
                throw;
            }
        }

        #endregion
    }
}