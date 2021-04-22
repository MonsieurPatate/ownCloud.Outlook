using System.ComponentModel;
using System.Runtime.CompilerServices;
using Newtonsoft.Json;
using ownCloud.Outlook.Annotations;

namespace ownCloud.Outlook
{
    public class Config : INotifyPropertyChanged
    {
        private string _server;

        /// <summary>
        /// Owncloud server url
        /// </summary>
        [JsonProperty("Server")]
        public string Server
        {
            get => _server;
            set
            {
                if (_server == value) return;
                _server = value;
                OnPropertyChanged();
            }

        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}