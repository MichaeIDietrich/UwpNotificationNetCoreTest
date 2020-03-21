using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace UwpNotificationNetCoreTest
{
    public static class Registration
    {
        private const int CLASS_E_NOAGGREGATION = -2147221232;
        private const int E_NOINTERFACE = -2147467262;
        private const int CLSCTX_LOCAL_SERVER = 4;
        private const int REGCLS_MULTIPLEUSE = 1;
        private const int S_OK = 0;

        private static readonly Guid IUnknownGuid = new Guid("00000000-0000-0000-C000-000000000046");

        private static uint _cookie;

        [DllImport("ole32.dll")]
        private static extern int CoRegisterClassObject(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            [MarshalAs(UnmanagedType.IUnknown)] object pUnk,
            uint dwClsContext,
            uint flags,
            out uint lpdwRegister);

        [DllImport("ole32.dll")]
        private static extern int CoRevokeClassObject(uint dwRegister);

        public static EventHandler<NotificationReceivedEventArgs> NotificationReceived;

        public static void RegisterApplication()
        {
            RegisterApplicationId();

            RegisterProtocolScheme();

            RegisterComServerInRegistry();

            var uuid = typeof(NotificationActivator).GUID;

            // register a class factory that is used by COM to create a new instance of out toast activation callback handler
            // not completely sure what happens when this code is called from multiple app instances
            CoRegisterClassObject(uuid, new NotificationActivatorClassFactory(), CLSCTX_LOCAL_SERVER,
                REGCLS_MULTIPLEUSE, out _cookie);
        }

        public static void UnregisterApplication()
        {
            // should be invoked on shutdown, since our class factory can no longer be used to create new instances.
            if (_cookie != 0)
                CoRevokeClassObject(_cookie);
        }

        public static void ClearAll()
        {
            UnregisterApplication();

            Registry.CurrentUser.DeleteSubKeyTree($"SOFTWARE\\Classes\\CLSID\\{{{Defines.ComServerGuid}}}");

            Registry.CurrentUser.DeleteSubKeyTree($@"SOFTWARE\Classes\{Defines.ProtocolScheme}");

            var userPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Microsoft\Windows\Start Menu\Programs", $"{Defines.AppId}.lnk");

            if (File.Exists(userPath))
                File.Delete(userPath);
        }

        private static void RegisterComServerInRegistry()
        {
            var exePath = Process.GetCurrentProcess().MainModule.FileName;

            // here we define that our app is to be launched when the toast notification is activated
            // and the app is not running
            var regString = $"SOFTWARE\\Classes\\CLSID\\{{{Defines.ComServerGuid}}}\\LocalServer32";

            using (var key = Registry.CurrentUser.OpenSubKey(regString))
            {
                if (string.Equals(key?.GetValue(null) as string, exePath, StringComparison.OrdinalIgnoreCase))
                    return;
            }

            using (var key = Registry.CurrentUser.CreateSubKey(regString))
            {
                key?.SetValue(null, exePath);
            }
        }

        private static void RegisterApplicationId()
        {
            // to make desktop notifications work for non packaged applications,
            // it seems to be necessary to have a shortcut that points to the application
            // registered with a valid application id which then can used by the ToastNotificationManager

            // and we also fill in the GUID of our COM server for toast activation 

            var comServerGuid = Guid.Parse(Defines.ComServerGuid);
            var userPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Microsoft\Windows\Start Menu\Programs", $"{Defines.AppId}.lnk");

            using var shortcut = new ShellLink();

            if (File.Exists(userPath))
                shortcut.Load(userPath);

            if (shortcut.AppUserModelID == Defines.AppId
                && shortcut.AppUserModelToastActivatorCLSID == comServerGuid)
                return;

            shortcut.TargetPath = Defines.ExecutablePath;
            shortcut.AppUserModelID = Defines.AppId;
            shortcut.AppUserModelToastActivatorCLSID = comServerGuid;

            shortcut.Save(userPath);
        }

        private static void RegisterProtocolScheme()
        {
            using var key = Registry.CurrentUser.CreateSubKey($@"SOFTWARE\Classes\{Defines.ProtocolScheme}");

            key.SetValue(null, $"URL:{Defines.ProtocolScheme}");
            key.SetValue("URL Protocol", "");

            using var commandKey = key.CreateSubKey(@"shell\open\command");

            commandKey.SetValue("",  $"\"{Defines.ExecutablePath}\" \"%1\"");
        }

        [ComImport]
        [Guid("00000001-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IClassFactory
        {
            [PreserveSig]
            int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject);

            [PreserveSig]
            int LockServer(bool fLock);
        }

        private class NotificationActivatorClassFactory : IClassFactory
        {
            public int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject)
            {
                ppvObject = IntPtr.Zero;

                if (pUnkOuter != IntPtr.Zero)
                    Marshal.ThrowExceptionForHR(CLASS_E_NOAGGREGATION);

                if (riid == typeof(NotificationActivator).GUID || riid == IUnknownGuid)
                    // Create the instance of the .NET object
                    ppvObject = Marshal.GetComInterfaceForObject(new NotificationActivator(),
                        typeof(INotificationActivationCallback));
                else
                    // The object that ppvObject points to does not support the
                    // interface identified by riid.
                    Marshal.ThrowExceptionForHR(E_NOINTERFACE);
                return S_OK;
            }

            public int LockServer(bool fLock)
            {
                return S_OK;
            }
        }

        public sealed class NotificationReceivedEventArgs : EventArgs
        {
            public NotificationReceivedEventArgs(string arguments, IReadOnlyDictionary<string, string> data)
            {
                Arguments = arguments;
                Data = data;
            }

            public string Arguments { get; }

            public IReadOnlyDictionary<string, string> Data { get; }
        }

        [ComImport]
        [Guid("53E31837-6600-4A81-9395-75CFFE746F94")]
        [ComVisible(true)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface INotificationActivationCallback
        {
            void Activate(
                [In] [MarshalAs(UnmanagedType.LPWStr)] string appUserModelId,
                [In] [MarshalAs(UnmanagedType.LPWStr)] string invokedArgs,
                [In] [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 3)]
                NOTIFICATION_USER_INPUT_DATA[] data,
                [In] [MarshalAs(UnmanagedType.U4)] uint dataCount);
        }


        [Serializable]
        public struct NOTIFICATION_USER_INPUT_DATA
        {
            [MarshalAs(UnmanagedType.LPWStr)] public readonly string Key;

            [MarshalAs(UnmanagedType.LPWStr)] public readonly string Value;
        }

        [ClassInterface(ClassInterfaceType.None)]
        [ComSourceInterfaces(typeof(INotificationActivationCallback))]
        [Guid(Defines.ComServerGuid)]
        [ComVisible(true)]
        public class NotificationActivator : INotificationActivationCallback
        {
            #region interface INotificationActivationCallback

            public void Activate(string appUserModelId, string invokedArgs, NOTIFICATION_USER_INPUT_DATA[] data,
                uint dataCount)
            {
                var keyValuePairs = Enumerable.Range(0, (int) dataCount)
                    .ToDictionary(i => data[i].Key, i => data[i].Value);

                NotificationReceived?.Invoke(this, new NotificationReceivedEventArgs(invokedArgs, keyValuePairs));
            }

            #endregion
        }
    }
}