using System;
using System.Linq;
using System.Windows;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;

namespace UwpNotificationNetCoreTest
{
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();

            Registration.NotificationReceived += OnNotificationReceivedFromCom;

            DataTransfer.RegisterCallback(this, OnNotificationReceivedFromProtocol);
        }

        private void OnNotificationReceivedFromCom(object sender, Registration.NotificationReceivedEventArgs e)
        {
            var arguments = $"{e.Arguments}{string.Join("", e.Data.Select(item => $"&{item.Key}={item.Value}"))}";

            AddNotificationEntry($"[COM] {arguments}");
        }

        private void OnNotificationReceivedFromProtocol(string arguments)
        {
            if (arguments.StartsWith(Defines.ProtocolScheme, StringComparison.OrdinalIgnoreCase))
                arguments = arguments.Substring(Defines.ProtocolScheme.Length + 1);

            AddNotificationEntry($"[PROTOCOL] {arguments}");
        }

        private void AddNotificationEntry(string content)
        {
            Dispatcher.Invoke(() =>
            {
                ListBox.Items.Add($"({DateTime.Now}) notification received > {content}");

                ForceFocus();
            });
        }

        private void ForceFocus()
        {
            if (!IsVisible)
                Show();

            if (WindowState == WindowState.Minimized)
                WindowState = WindowState.Normal;

            Activate();
            Topmost = true;
            Topmost = false;
            Focus();
        }

        private void OnCreateNotificationForCom(object sender, RoutedEventArgs e)
        {
            var xml = @"<toast launch=""arguments"" >
<visual>
<binding template=""ToastGeneric"">
<image placement=""appLogoOverride"" src="""" />
<text>Header</text>
<text>Category</text>
<group>
<subgroup>
<text hint-maxLines=""2"" hint-wrap=""true"">Message</text>
</subgroup>
</group>
<text placement=""attribution"" />
</binding>
</visual>
<actions><action content=""OK"" arguments=""ok"" /><action content=""Cancel"" arguments=""cancel"" /></actions>
</toast>";
            var document = new XmlDocument();
            document.LoadXml(xml);

            var toast = new ToastNotification(document);

            ToastNotificationManager.CreateToastNotifier(Defines.AppId).Show(toast);
        }

        private void OnCreateNotificationForProtocol(object sender, RoutedEventArgs e)
        {
            var xml = $@"<toast activationType=""protocol"" launch=""{Defines.ProtocolScheme}:arguments"" >
<visual>
<binding template=""ToastGeneric"">
<image placement=""appLogoOverride"" src="""" />
<text>Header</text>
<text>Category</text>
<group>
<subgroup>
<text hint-maxLines=""2"" hint-wrap=""true"">Message</text>
</subgroup>
</group>
<text placement=""attribution"" />
</binding>
</visual>
<actions>
<action content=""OK"" arguments=""{Defines.ProtocolScheme}:ok"" activationType=""protocol"" />
<action content=""Cancel"" arguments=""{Defines.ProtocolScheme}:cancel"" activationType=""protocol"" />
</actions>
</toast>";
            var document = new XmlDocument();
            document.LoadXml(xml);

            var toast = new ToastNotification(document);

            ToastNotificationManager.CreateToastNotifier(Defines.AppId).Show(toast);
        }
    }
}