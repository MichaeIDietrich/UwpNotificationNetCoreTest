using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace UwpNotificationNetCoreTest
{
    public static class DataTransfer
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref CopyData target);

        private const int WM_COPYDATA = 0x4A;
        private const int MessageId = 123;

        [StructLayout(LayoutKind.Sequential)]
        private struct CopyData
        {
            public IntPtr dwData;
            public int cbData;
            public IntPtr lpData;
        }

        public static void Send(IntPtr targetHandle, string value)
        {
            var data = new CopyData
            {
                dwData = (IntPtr)MessageId,
                lpData = Marshal.StringToCoTaskMemUni(value),
                cbData = 2 * value.Length + 1
            };

            SendMessage(targetHandle, WM_COPYDATA, IntPtr.Zero, ref data);

            Marshal.FreeCoTaskMem(data.lpData);
            data.lpData = IntPtr.Zero;
            data.cbData = 0;
        }

        public static void RegisterCallback(Window window, Action<string> callback)
        {
            IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam, ref bool handled)
            {
                if (msg == WM_COPYDATA)
                {
                    var data = Marshal.PtrToStructure<CopyData>(lparam);

                    if (data.dwData == (IntPtr) MessageId)
                    {
                        var value = Marshal.PtrToStringUni(data.lpData);

                        callback(value);
                        handled = true;
                    }
                }

                return IntPtr.Zero;
            }

            var source = (HwndSource)PresentationSource.FromDependencyObject(window);

            if (source is null)
            {
                window.SourceInitialized += delegate { RegisterCallback(window, callback); };
                return;
            }

            source.AddHook(WndProc);
        }

    }
}
