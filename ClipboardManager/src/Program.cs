using System;
using System.Windows.Forms;

namespace ClipboardManager {
    internal static class Program {
        private static ClipboardApplication clipboardUI;

        [STAThread]
        private static void Main() {
            try {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                if (SingleInstanceApplicationHandler.IsAppRunning())
                    throw new InvalidOperationException("已經有另一個啟動了。");
                SingleInstanceApplicationHandler.StartService();
                Application.Idle += Run;
                Application.Run();
                SingleInstanceApplicationHandler.StopService();
            } catch (Exception ex) {
                UIHelper.ShowErrorMessage(ex);
            }
        }

        private static void Run(object sender, EventArgs e) {
            Application.Idle -= Run;
            clipboardUI = new ClipboardApplication();
        }
    }
}
