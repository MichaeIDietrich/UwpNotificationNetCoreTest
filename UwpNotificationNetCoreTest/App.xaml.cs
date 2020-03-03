using System.Windows;

namespace UwpNotificationNetCoreTest
{
    public partial class App
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            Registration.RegisterApplication();

            base.OnStartup(e);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Registration.UnregisterApplication();

            base.OnExit(e);
        }
    }
}