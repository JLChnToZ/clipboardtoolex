using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using TSItem = System.Windows.Forms.ToolStripItem;
using TSItems = System.Windows.Forms.ToolStripItemCollection;
using TSLabel = System.Windows.Forms.ToolStripLabel;
using TSMenuItem = System.Windows.Forms.ToolStripMenuItem;
using TSSeparator = System.Windows.Forms.ToolStripSeparator;

namespace ClipboardManager {
    internal class HistoryMenu {
        public bool removeOnUse;
        public readonly TSItems root;
        public readonly TSItem index;
        private readonly TSMenuItem clearMenu;
        private readonly List<TSMenuItem> list = new List<TSMenuItem>();
        private readonly HashSet<TSMenuItem> unused = new HashSet<TSMenuItem>();

        public HistoryMenu(TSItems root, string text) {
            this.root = root;
            root.AddRange(new TSItem[] {
                index = new TSLabel(text),
                clearMenu = new TSMenuItem(Language.ClearHistory, null, HandleClearClick) {
                    Enabled = false,
                },
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
            root.Insert(root.IndexOf(index) + 1, item);
            list.Add(item);
            UpdateClearEnabled();
        }

        public void ClearHistory(int howManyLeft = 0) {
            while (list.Count > howManyLeft)
                RemoveHistory(list[0]);
            UpdateClearEnabled();
        }

        private void RemoveHistory(TSMenuItem item) {
            root.Remove(item);
            list.Remove(item);
            unused.Add(item);
            item.Tag = null;
            foreach (TSItem child in item.DropDownItems) {
                if (child.BackgroundImage != null) {
                    child.BackgroundImage.Dispose();
                    child.BackgroundImage = null;
                }
                if (child.Image != null) {
                    child.Image.Dispose();
                    child.Image = null;
                }
            }
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
            UpdateClearEnabled();
        }

        private void HandleClearClick(object sender, EventArgs e) {
            ClearHistory();
            UpdateClearEnabled();
        }

        private void UpdateClearEnabled() {
            clearMenu.Enabled = list.Count > 0;
        }

        private static DataObject CloneDataObject(IDataObject dataObject) {
            DataObject result = new DataObject();
            foreach (string format in dataObject.GetFormats(false))
                result.SetData(format, dataObject.GetData(format, false));
            return result;
        }
    }
}
