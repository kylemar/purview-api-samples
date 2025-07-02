using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Purview_API_Explorer
{
    public enum ProtectionScopeMessageType
    {
        NoState,
        Modified,
        TooSoon
    }
    /// <summary>
    /// Interaction logic for ProtectionScopeInfo.xaml
    /// </summary>
    public partial class ProtectionScopeInfo : Window
    {
        public ProtectionScopeInfo(ProtectionScopeMessageType type)
        {
            InitializeComponent();

            if (type == ProtectionScopeMessageType.NoState)
            {
                TitleText.Text = "You need to call ProtectionScopes/compute."; 
                MessageText.Text =
                    "You may be calling ProcessContent, or awaiting results, when not needed.\r\n\r\n" +
                    "Do not call ProtectionScopes/compute call every time you call ProcessContent\r\n" +
                    "Call ProtectionScopes/compute\r\n" +
                    "\tWhen the user first starts the app\r\n" +
                    "\tAfter ProcessContent returns\r\n\t\tprotectionScopeState as 'modified'\r\n" +
                    "\tAfter 30 minuites of idle time.";
            }
            else if (type == ProtectionScopeMessageType.Modified)
            {
                TitleText.Text = "Protection Scope State Modified.";
                MessageText.Text = "Call ProtectionScopes/compute to get the latest state.\r\n\r\n" +
                                   "ProcessContent will also return protectionScopeState as 'modified' if you did not previously call ProtectionScopes/compute or set the 'If-None-Match' header\r\n\r\n" +
                                   "Do not call ProtectionScopes/compute call every time you call ProcessContent.\r\n";
            }
            else if (type == ProtectionScopeMessageType.TooSoon)
            {
                TitleText.Text = "Don't call ProtectionScopes/compute too often";
                MessageText.Text = "Use ProtectionScopes/compute to optimize your integration with Microsoft Purview.\r\n" +
                    "Do not call ProtectionScopes/compute call every time you call ProcessContent\r\n" +
                    "Call ProtectionScopes/compute\r\n" +
                    "\tWhen the user first starts the app\r\n" +
                    "\tAfter ProcessContent returns\r\n\t\tprotectionScopeState as 'modified'\r\n" +
                    "\tAfter 30 minuites of idle time.";

            }
            else
            {
                MessageText.Text = "Unknown scope type.";
            }

        }

        private void DocumentationButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://learn.microsoft.com/en-us/purview/developer/use-the-api#compute-protection-scopes") { UseShellExecute = true });
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
