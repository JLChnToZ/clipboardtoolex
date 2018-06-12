using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace ClipboardManager {
    internal static class Native {
        public const uint SHGFI_ICON = 0x100;
        public const uint SHGFI_SMALLICON = 0x1;
        public const int WM_DRAWCLIPBOARD = 0x308;
        public const int WM_CHANGECBCHAIN = 0x030D;

        [StructLayout(LayoutKind.Sequential)]
        public struct SHFILEINFO {
            public IntPtr handle;
            public IntPtr index;
            public uint attr;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string display;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string type;
        };

        [DllImport("shell32.dll")]
        public static extern IntPtr SHGetFileInfo(string path, uint fattrs, ref SHFILEINFO sfi, uint size, uint flags);
        
        [DllImport("User32.dll")]
        public static extern int SetClipboardViewer(int hWndNewViewer);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);

        public static Icon GetSmallIcon(string path) {
            SHFILEINFO info = new SHFILEINFO();
            SHGetFileInfo(path, 0, ref info, (uint)Marshal.SizeOf(info), SHGFI_ICON | SHGFI_SMALLICON);
            return Icon.FromHandle(info.handle);
        }
    }
}
