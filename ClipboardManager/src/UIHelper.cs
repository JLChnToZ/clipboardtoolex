using System;
using System.Windows.Forms;

namespace ClipboardManager {
    public static class UIHelper {

        public static void ShowErrorMessage(Exception ex) {
            MessageBox.Show(
                ex.Message,
                Application.ProductName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1
            );
        }
    }
}
