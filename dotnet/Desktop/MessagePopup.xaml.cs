using System.Windows;

namespace PurviewAPIExp
{
    public partial class MessagePopup : Window
    {
        public MessagePopup(string message, string title = "Message")
        {
            InitializeComponent();
            TitleText.Text = title;
            MessageText.Text = message;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
    }
}