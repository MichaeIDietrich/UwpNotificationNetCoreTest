using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace UwpNotificationNetCoreTest
{
    public partial class App
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            if (SignalToPrimaryInstance())
            {
                Shutdown();
                return;
            }

            Registration.RegisterApplication();

            base.OnStartup(e);

            new MainWindow().Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Registration.UnregisterApplication();

            base.OnExit(e);
        }

        private static bool SignalToPrimaryInstance()
        {
            var commandLine = string.Join("", Environment.GetCommandLineArgs().Skip(1));

            if (!commandLine.StartsWith(Defines.ProtocolScheme, StringComparison.OrdinalIgnoreCase))
                return false;
            
            var mainProcess = Process
                .GetProcesses()
                .FirstOrDefault(IsPrimaryInstance);

            if (mainProcess is null)
            {
                SendToOwnMainWindowWithDelay(commandLine);
                return false;
            }

            DataTransfer.Send( mainProcess.MainWindowHandle, commandLine);

            return true;
        }

        private static bool IsPrimaryInstance(Process process)
        {
            try
            {
                return process.MainModule?.FileName == Defines.ExecutablePath &&
                       process.MainWindowHandle != IntPtr.Zero;
            }
            catch
            {
                // don't know whether there is a better way to search without try/catch
                return false;
            }
        }

        private static async void SendToOwnMainWindowWithDelay(string value)
        {
            // when this is the primary instance and activated by protocol
            // we use a short delay to send the message the main window

            await Task.Delay(1000);

            var handle = new WindowInteropHelper(Current.MainWindow).Handle;

            DataTransfer.Send(handle, value);

        }
    }
}