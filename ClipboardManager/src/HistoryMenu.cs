using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using TSItem = System.Windows.Forms.ToolStripItem;
using TSMenuItem = System.Windows.Forms.ToolStripMenuItem;
using TSSeparator = System.Windows.Forms.ToolStripSeparator;

namespace ClipboardManager {
    internal class HistoryMenu {
        public bool removeOnUse;
        public readonly TSMenuItem root;
        private readonly List<TSMenuItem> list = new List<TSMenuItem>();
        private readonly HashSet<TSMenuItem> unused = new HashSet<TSMenuItem>();

        public HistoryMenu(string text) {
            root = new TSMenuItem(text);
            root.DropDownItems.AddRange(new TSItem[] {
                new TSSeparator(),
                new TSMenuItem(Language.ClearHistory, null, HandleClearClick),
            });
        }

        public void RecordHistory() {
            RecordHistory(Clipboard.GetDataObject());
        }

        public void RecordHistory(IDataObject dataObject) {
            TSMenuItem item;
            if (unused.Count > 0) {
                item = unused.First();
                unused.Remove(item);
            } else {
                item = new TSMenuItem();
                item.DropDownItems.AddRange(new TSItem[] {
                    new TSMenuItem(Language.RestoreHistory, null, HandleUseClick) {
                        Tag = item
                    },
                    new TSMenuItem(Language.ForgetHistory, null, HandleRemoveClick) {
                        Tag = item
                    },
                    new TSSeparator(),
                });
            }
            ClipboardApplication.UpdateDisplay(dataObject, item);
            item.Tag = CloneDataObject(dataObject);
            item.Enabled = true;
            root.DropDownItems.Insert(list.Count, item);
            list.Add(item);
        }

        public void ClearHistory(int howManyLeft = 0) {
            while (list.Count > howManyLeft)
                RemoveHistory(list[0]);
        }

        private void RemoveHistory(TSMenuItem item) {
            item.Tag = null;
            root.DropDownItems.Remove(item);
            list.RemoveAt(0);
            unused.Add(item);
        }

        private void HandleUseClick(object sender, EventArgs e) {
            TSMenuItem item = (sender as TSMenuItem).Tag as TSMenuItem;
            IDataObject dataObject = item.Tag as IDataObject;
            if (removeOnUse) RemoveHistory(item);
            try {
                Clipboard.SetDataObject(dataObject, true);
            } catch (Exception ex) {
                UIHelper.ShowErrorMessage(ex);
            }
        }

        private void HandleRemoveClick(object sender, EventArgs e) {
            RemoveHistory((sender as TSMenuItem).Tag as TSMenuItem);
        }

        private void HandleClearClick(object sender, EventArgs e) {
            ClearHistory();
        }

        private static DataObject CloneDataObject(IDataObject dataObject) {
            DataObject result = new DataObject();
            foreach (string format in dataObject.GetFormats(false))
                result.SetData(format, dataObject.GetData(format, false));
            return result;
        }
    }
}
