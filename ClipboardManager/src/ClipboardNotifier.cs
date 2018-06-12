using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ClipboardManager {
    [DefaultEvent("ClipboardChanged")]
    public partial class ClipboardNotifier: Control {
        IntPtr nextClipboardViewer;

        public ClipboardNotifier() {
            Visible = false;
            nextClipboardViewer = (IntPtr)Native.SetClipboardViewer((int)Handle);
        }
        
        public event EventHandler<ClipboardChangedEventArgs> ClipboardChanged;

        protected override void Dispose(bool disposing) {
            try {
                Native.ChangeClipboardChain(Handle, nextClipboardViewer);
            } catch { }
        }

        protected override void WndProc(ref Message message) {
            switch (message.Msg) {
                case Native.WM_DRAWCLIPBOARD:
                    OnClipboardChanged();
                    Native.SendMessage(nextClipboardViewer, message.Msg, message.WParam, message.LParam);
                    break;

                case Native.WM_CHANGECBCHAIN:
                    if (message.WParam == nextClipboardViewer)
                        nextClipboardViewer = message.LParam;
                    else
                        Native.SendMessage(nextClipboardViewer, message.Msg, message.WParam, message.LParam);
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
