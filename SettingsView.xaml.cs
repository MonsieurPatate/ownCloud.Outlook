using System.Windows;
using System.Windows.Input;
using AdysTech.CredentialManager;

namespace ownCloud.Outlook
{
    /// <summary>
    /// Interaction logic for SettingsView.xaml
    /// </summary>
    public partial class SettingsView : Window
    {
        private Config Config => (Config) DataContext;

        public SettingsView()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DataContext = ConfigManager.Read();
        }

        private void EnterAuthorization_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var defaultUserName = string.Empty;
            var credential = CredentialManager.GetCredentials(Constants.AddInName);
            if (credential != null)
                defaultUserName = credential.UserName;

            var save = true;
            credential = CredentialManager.PromptForCredentials(Constants.AddInName, ref save, $"Please, enter your credentials for {Config.Server}", "Credentials for ownCloud.Outlook AddIn", defaultUserName);
            if (credential != null)
                CredentialManager.SaveCredentials(Constants.AddInName, credential);
        }

        //private void EnterAuthorizationCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        //{
        //    e.CanExecute = true;
        //}

        private void SaveCommand_Execute(object sender, ExecutedRoutedEventArgs e)
        {
            ConfigManager.Save(Config);
        }
    }
}
