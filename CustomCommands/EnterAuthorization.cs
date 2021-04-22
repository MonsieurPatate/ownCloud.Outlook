using System.Windows.Input;

namespace ownCloud.Outlook.CustomCommands
{
    public class EnterAuthorization : RoutedUICommand
    {
        public EnterAuthorization()
            : base("EnterAuthorization", "EnterAuthorization", typeof(EnterAuthorization))
        {
        }
    }
}