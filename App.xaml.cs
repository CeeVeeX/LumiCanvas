using Microsoft.UI.Xaml;
using Windows.ApplicationModel.Activation;

namespace LumiCanvas
{
    public partial class App : Application
    {
        private MainWindow? _window;
        private readonly WorkspaceSession _session = new();

        public App()
        {
            InitializeComponent();
        }

        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            _window = new MainWindow(_session);
            _window.Activate();
            _window.HideToBackground();
        }
    }
}
