using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ClipboardManager {
    [DefaultEvent("ClipboardChanged")]
    public partial class ClipboardNotifier: Control {
        const int WM_DRAWCLIPBOARD = 0x308;
        const int WM_CHANGECBCHAIN = 0x030D;

        [DllImport("User32.dll")]
        protected static extern int SetClipboardViewer(int hWndNewViewer);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        protected static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        protected static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);

        IntPtr nextClipboardViewer;

        public ClipboardNotifier() {
            Visible = false;
            nextClipboardViewer = (IntPtr)SetClipboardViewer((int)Handle);
        }
        
        public event EventHandler<ClipboardChangedEventArgs> ClipboardChanged;

        protected override void Dispose(bool disposing) {
            try {
                ChangeClipboardChain(Handle, nextClipboardViewer);
            } catch { }
        }

        protected override void WndProc(ref Message message) {
            switch (message.Msg) {
                case WM_DRAWCLIPBOARD:
                    OnClipboardChanged();
                    SendMessage(nextClipboardViewer, message.Msg, message.WParam, message.LParam);
                    break;

                case WM_CHANGECBCHAIN:
                    if (message.WParam == nextClipboardViewer)
                        nextClipboardViewer = message.LParam;
                    else
                        SendMessage(nextClipboardViewer, message.Msg, message.WParam, message.LParam);
                    break;

                default:
                    base.WndProc(ref message);
                    break;
            }
        }

        private void OnClipboardChanged() {
            IDataObject data = Clipboard.GetDataObject();
            ClipboardChanged?.Invoke(this, new ClipboardChangedEventArgs(data));
        }
    }

    public class ClipboardChangedEventArgs: EventArgs {
        public readonly IDataObject dataObject;

        public ClipboardChangedEventArgs(IDataObject dataObject) {
            this.dataObject = dataObject;
        }
    }
}
