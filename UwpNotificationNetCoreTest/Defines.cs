using System.Diagnostics;

namespace UwpNotificationNetCoreTest
{
    public static class Defines
    { 
        // application id that needs to be unique for every application
        public const string AppId = "UwpNotificationNetCoreTest";

        // COM server GUID that needs to be unique for every application
        public const string ComServerGuid = "9DDCD0D6-6B91-4245-B76E-03EEF2C39998";

        public const string ProtocolScheme = "UwpNotificationNetCoreTest";

        public static readonly string ExecutablePath = Process.GetCurrentProcess().MainModule.FileName;
    }
}
