using System;
using System.Runtime.InteropServices;

namespace RFIDDemoCS
{
    public class SysIoApi
    {
        public const int LEFT_TRIGGER_KEY = 1;
        public const int RIGHT_TRIGGER_KEY = 2;

        [DllImport("sysioapi.dll")]
        public static extern bool TriggerKeyStatus(int key);
    }

    public class CoreDLL
    {
        public const string TriggerKeyEvent = "KeybdTriggerChangeEvent";

        [DllImport("CoreDLL.dll")]
        public static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState, string lpName);

        [DllImport("CoreDLL.dll")]
        public static extern bool CloseHandle(IntPtr hObject);
    }
}
