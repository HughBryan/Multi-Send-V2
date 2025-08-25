using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Multi_Send
{
    public class RecipientViewModel : INotifyPropertyChanged
    {
        private string email = "";
        private string name = "";

        public string Email
        {
            get => email;
            set { email = value; OnPropertyChanged(); }
        }

        public string Name
        {
            get => name;
            set { name = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
