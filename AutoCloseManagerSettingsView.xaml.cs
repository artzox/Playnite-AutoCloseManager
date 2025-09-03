using Playnite.SDK;
using System.Windows.Controls;

namespace AutoCloseManagerPlugin
{
    public partial class AutoCloseManagerSettingsView : UserControl
    {
        public AutoCloseManagerSettingsView()
        {
            InitializeComponent();
        }

        public AutoCloseManagerSettingsView(AutoCloseManagerSettings settings)
        {
            InitializeComponent();
            DataContext = settings;
        }
    }
}