using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using ClipboardManager.Properties;

namespace ClipboardManager {
    internal class ClipboardApplication {
        private readonly Settings settings = Settings.Default;
        private readonly NotifyIcon notifyIcon = new NotifyIcon {
            Icon = SystemIcons.Application,
            Text = Application.ProductName,
            ContextMenuStrip = new ContextMenuStrip(),
            Visible = true,
        };
        private readonly Timer cleanupTimer = new Timer {
            Enabled = false,
        };
        private readonly ClipboardNotifier clipboardNotifier = new ClipboardNotifier();
        private readonly ToolStripMenuItem dataDisplay = new ToolStripMenuItem {
            Visible = false,
        };
        private readonly ToolStripLabel imageDataDisplay = new ToolStripLabel {
            Visible = false,
            AutoSize = false,
            BackgroundImageLayout = ImageLayout.Zoom,
        };
        private readonly ToolStripMenuItem pinnedList = new ToolStripMenuItem("已手動記住 (&I)");
        private readonly ToolStripMenuItem historyList = new ToolStripMenuItem("歷史記錄 (&H)");

        private readonly List<ToolStripMenuItem>
            pinnedItems = new List<ToolStripMenuItem>(),
            historyItems = new List<ToolStripMenuItem>();
        private readonly HashSet<ToolStripMenuItem>
            unusedPinnedItems = new HashSet<ToolStripMenuItem>(),
            unusedHistoryItems = new HashSet<ToolStripMenuItem>();

        private const string RUN_KEY = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
        private static bool RunAtStartup {
            get {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RUN_KEY, false)) {
                    return key.GetValue(Application.ProductName) is string value &&
                        string.Equals(value, Application.ExecutablePath, StringComparison.Ordinal);
                }
            }
            set {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RUN_KEY, true)) {
                    if (value)
                        key.SetValue(Application.ProductName, Application.ExecutablePath);
                    else
                        key.DeleteValue(Application.ProductName);
                }
            }
        }

        public ClipboardApplication() {
            clipboardNotifier.ClipboardChanged += HandleClipboard;
            
            if(settings.autoCleanInterval > 0)
                cleanupTimer.Interval = settings.autoCleanInterval * 1000;
            cleanupTimer.Tick += HandleCleanUpClicked;
            
            notifyIcon.DoubleClick += HandleCleanUpClicked;

            NumericUpDown intervalItemBox = new NumericUpDown {
                Minimum = 0,
                Maximum = int.MaxValue,
                Value = settings.autoCleanInterval,
            };
            intervalItemBox.ValueChanged += HandleIntervalChange;
            NumericUpDown historyCountBox = new NumericUpDown {
                Minimum = 0,
                Maximum = int.MaxValue,
                Value = settings.maxHistoryObjects,
            };
            historyCountBox.ValueChanged += HandleHistoryCountChange;

            notifyIcon.ContextMenuStrip.Items.AddRange(new ToolStripItem[] {
                new ToolStripLabel(string.Format("{0} - 版本 {1}", Application.ProductName, Application.ProductVersion)),

                new ToolStripSeparator(),

                new ToolStripLabel("剪貼簿現有資料"),
                dataDisplay, imageDataDisplay,
                new ToolStripMenuItem("立即清理 (&C)", null, HandleCleanUpClicked),

                new ToolStripSeparator(),

                new ToolStripLabel("自動在…秒後清理 (設為 0 則不自動清理) (&A)"),
                new ToolStripControlHost(intervalItemBox),
                new ToolStripLabel("自動記住…次記錄 (&R)"),
                new ToolStripControlHost(historyCountBox),

                new ToolStripSeparator(),

                pinnedList, historyList,

                new ToolStripSeparator(),

                new ToolStripMenuItem("提示剪貼簿被更改 (&N)", null, HandleNotifyToggle) {
                    Checked = settings.notifyEnabled,
                },
                new ToolStripMenuItem("開機時啟動 (&S)", null, RunAtStartupClick) {
                    Checked = RunAtStartup
                },

                new ToolStripSeparator(),

                new ToolStripMenuItem("離開 (&E)", null, HandleExitClick),
            });

            pinnedList.DropDownItems.AddRange(new ToolStripItem[] {
                new ToolStripSeparator(),
                new ToolStripMenuItem("清理 (&C)", null, HandleClearPinnedClick),
            });

            historyList.Enabled = settings.maxHistoryObjects > 0;
            historyList.DropDownItems.AddRange(new ToolStripItem[] {
                new ToolStripSeparator(),
                new ToolStripMenuItem("清理 (&C)", null, HandleClearHistoryClick),
            });
            
            HandleClipboard(false, Clipboard.GetDataObject());
            Application.ApplicationExit += HandleApplicationExit;
        }

        #region Event Handlers
        private void HandleClipboard(bool notify, IDataObject dataObject, bool startTimer = true) {
            bool hasData = UpdateDisplay(dataObject, dataDisplay, imageDataDisplay);
            if (notify) NotifyData(dataObject);
            if (hasData) {
                if (startTimer && settings.autoCleanInterval > 0)
                    cleanupTimer.Start();
                if (historyList.HasDropDownItems) {
                    ClearHistory(historyList, historyItems, unusedHistoryItems);
                    if (settings.maxHistoryObjects > 0)
                        RecordHistory(dataObject, historyList, historyItems, unusedHistoryItems, HandleUseHistoryClick, HandleRemoveHistoryClick);
                }
            }
        }

        private void HandleIntervalChange(object sender, EventArgs e) {
            NumericUpDown interval = sender as NumericUpDown;
            cleanupTimer.Stop();
            settings.autoCleanInterval = (int)interval.Value;
            if (interval.Value > 0) {
                cleanupTimer.Interval = settings.autoCleanInterval * 1000;
                HandleClipboard(false, Clipboard.GetDataObject());
            }
            settings.Save();
        }

        private void HandleHistoryCountChange(object sender, EventArgs e) {
            settings.maxHistoryObjects = (int)(sender as NumericUpDown).Value;
            ClearHistory(historyList, historyItems, unusedHistoryItems);
            historyList.Enabled = settings.maxHistoryObjects > 0;
            settings.Save();
        }

        private void HandleNotifyToggle(object sender, EventArgs e) {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            item.Checked = settings.notifyEnabled = !item.Checked;
            settings.Save();
        }

        private void HandleClipboard(object sender, ClipboardChangedEventArgs e) {
            cleanupTimer.Stop();
            HandleClipboard(settings.notifyEnabled, e.dataObject);
        }

        private static void HandleCleanUpClicked(object sender, EventArgs e) {
            try {
                Clipboard.Clear();
            } catch (Exception ex) {
                UIHelper.ShowErrorMessage(ex);
            }
        }

        private static void HandleExitClick(object sender, EventArgs e) {
            Application.Exit();
        }

        private void HandlePinClick(object sender, EventArgs e) {
            RecordHistory(Clipboard.GetDataObject(), pinnedList, pinnedItems, unusedPinnedItems, HandleUsePinnedClick, HandleRemovePinnedClick);
        }

        private static void HandleUsePinnedClick(object sender, EventArgs e) {
            ToolStripMenuItem item = (sender as ToolStripMenuItem).Tag as ToolStripMenuItem;
            IDataObject dataObject = item.Tag as IDataObject;
            try {
                Clipboard.SetDataObject(dataObject, true);
            } catch (Exception ex) {
                UIHelper.ShowErrorMessage(ex);
            }
        }

        private void HandleRemovePinnedClick(object sender, EventArgs e) {
            RemoveHistory((sender as ToolStripMenuItem).Tag as ToolStripMenuItem, pinnedList, pinnedItems, unusedPinnedItems);
        }

        private void HandleClearPinnedClick(object sender, EventArgs e) {
            ClearHistory(pinnedList, pinnedItems, unusedPinnedItems, true);
        }

        private void HandleUseHistoryClick(object sender, EventArgs e) {
            ToolStripMenuItem item = (sender as ToolStripMenuItem).Tag as ToolStripMenuItem;
            IDataObject dataObject = item.Tag as IDataObject;
            RemoveHistory(item, historyList, historyItems, unusedHistoryItems);
            try {
                Clipboard.SetDataObject(dataObject, true);
            } catch (Exception ex) {
                UIHelper.ShowErrorMessage(ex);
            }
        }

        private void HandleRemoveHistoryClick(object sender, EventArgs e) {
            RemoveHistory((sender as ToolStripMenuItem).Tag as ToolStripMenuItem, historyList, historyItems, unusedHistoryItems);
        }

        private void HandleClearHistoryClick(object sender, EventArgs e) {
            ClearHistory(historyList, historyItems, unusedHistoryItems, true);
        }

        private static void HandleFileClick(object sender, EventArgs e) {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            try {
                if (item.Tag != null)
                    Process.Start(new ProcessStartInfo(item.Tag as string) {
                        UseShellExecute = true
                    });
            } catch (Exception ex) {
                UIHelper.ShowErrorMessage(ex);
            }
        }

        private static void RunAtStartupClick(object sender, EventArgs e) {
            try {
                ToolStripMenuItem item = sender as ToolStripMenuItem;
                item.Checked = RunAtStartup = !item.Checked;
            } catch (Exception ex) {
                UIHelper.ShowErrorMessage(ex);
            }
        }

        private void HandleApplicationExit(object sender, EventArgs e) {
            notifyIcon.Visible = false;
        }
        #endregion

        private static void RemoveHistory(ToolStripMenuItem item, ToolStripMenuItem parent, List<ToolStripMenuItem> list, HashSet<ToolStripMenuItem> recycle) {
            item.Tag = null;
            parent.DropDownItems.Remove(item);
            list.RemoveAt(0);
            recycle.Add(item);
        }

        private void ClearHistory(ToolStripMenuItem parent, List<ToolStripMenuItem> list, HashSet<ToolStripMenuItem> recycle, bool cleanAll = false) {
            ToolStripItemCollection children = parent.DropDownItems;
            int cleanCount = cleanAll ? 0 : Math.Max(0, settings.maxHistoryObjects - 1);
            while (list.Count > cleanCount)
                RemoveHistory(list[0], parent, list, recycle);
        }

        private static bool UpdateDisplay(IDataObject dataObject, ToolStripMenuItem dataDisplay, ToolStripLabel imageDataDisplay = null) {
            bool hasData = false;
            try {
                hasData = true;
                dataDisplay.Visible = true;
                dataDisplay.Enabled = false;
                dataDisplay.Text = string.Empty;
                if (dataDisplay.HasDropDownItems) {
                    if (imageDataDisplay == null) {
                        HashSet<ToolStripItem> pendingRemoveItems = new HashSet<ToolStripItem>();
                        foreach (ToolStripItem item in dataDisplay.DropDownItems)
                            if (item is ToolStripLabel) {
                                imageDataDisplay = item as ToolStripLabel;
                                break;
                            } else if (item.Tag is string)
                                pendingRemoveItems.Add(item);
                        foreach (ToolStripItem item in pendingRemoveItems)
                            dataDisplay.DropDownItems.Remove(item);
                    } else
                        dataDisplay.DropDownItems.Clear();
                }
                if (imageDataDisplay != null) {
                    imageDataDisplay.Visible = false;
                    imageDataDisplay.BackgroundImage = null;
                }
                if (dataObject.GetDataPresent(DataFormats.UnicodeText, true)) {
                    string data = dataObject.GetData(DataFormats.UnicodeText, true) as string;
                    data = data.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
                    if (data.Length > 256)
                        dataDisplay.Text = string.Format("{0}...", data.Substring(0, 30));
                    else
                        dataDisplay.Text = data;
                } else if (dataObject.GetDataPresent(DataFormats.Bitmap)) {
                    Image data = dataObject.GetData(DataFormats.Bitmap) as Image;
                    string imgMeta = string.Format("圖片 ({0}x{1})", data.Width, data.Height);
                    if (imageDataDisplay == null) {
                        imageDataDisplay = new ToolStripLabel {
                            Visible = false,
                            AutoSize = false,
                            BackgroundImageLayout = ImageLayout.Zoom
                        };
                        dataDisplay.DropDownItems.Add(imageDataDisplay);
                    }
                    imageDataDisplay.Visible = true;
                    imageDataDisplay.BackgroundImage = data;
                    imageDataDisplay.Width = Math.Min(250, Math.Max(data.Width, data.Height));
                    imageDataDisplay.Height = data.Width > data.Height ?
                        (int)((double)imageDataDisplay.Width * data.Height / data.Width) : imageDataDisplay.Width;
                    dataDisplay.Text = imgMeta;
                } else if (dataObject.GetDataPresent(DataFormats.FileDrop)) {
                    string[] data = dataObject.GetData(DataFormats.FileDrop) as string[];
                    dataDisplay.Text = string.Format("{0} 個檔案", data.Length);
                    dataDisplay.Enabled = true;
                    foreach (string fileName in data) {
                        Bitmap icon = null;
                        try {
                            icon = Icon.ExtractAssociatedIcon(fileName).ToBitmap();
                        } catch (Exception) { }
                        ToolStripMenuItem item = new ToolStripMenuItem(Path.GetFileName(fileName)) {
                            Tag = fileName,
                            Image = icon,
                        };
                        item.Click += HandleFileClick;
                        dataDisplay.DropDownItems.Add(item);
                    }
                } else if (dataObject.GetDataPresent(DataFormats.WaveAudio)) {
                    using (Stream data = dataObject.GetData(DataFormats.WaveAudio) as Stream)
                        dataDisplay.Text = string.Format("聲音 ({0} 位元組)", data.Length);
                } else {
                    dataDisplay.Visible = false;
                    hasData = false;
                }
            } catch (Exception) { }
            return hasData;
        }

        private void NotifyData(IDataObject dataObject) {
            if (!settings.notifyEnabled) return;
            try {
                notifyIcon.BalloonTipTitle = "剪貼簿已被修改";
                if (dataObject.GetDataPresent(DataFormats.UnicodeText, true)) {
                    string data = dataObject.GetData(DataFormats.UnicodeText, true) as string;
                    data = data.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
                    if (data.Length > 256)
                        notifyIcon.BalloonTipText = string.Format("{0}...(以下咯 {1} 個字)", data.Substring(0, 256), data.Length - 256);
                    else
                        notifyIcon.BalloonTipText = data;
                } else if (dataObject.GetDataPresent(DataFormats.Bitmap)) {
                    Image data = dataObject.GetData(DataFormats.Bitmap) as Image;
                    string imgMeta = string.Format("圖片資料 ({0}x{1})", data.Width, data.Height);
                    notifyIcon.BalloonTipText = imgMeta;
                } else if (dataObject.GetDataPresent(DataFormats.FileDrop)) {
                    string[] data = dataObject.GetData(DataFormats.FileDrop) as string[];
                    StringBuilder sb = new StringBuilder();
                    if (data.Length < 3 && data.Length > 0) {
                        foreach (string fileName in data) {
                            if (sb.Length > 0) sb.Append(", ");
                            sb.Append(Path.GetFileName(fileName));
                        }
                    } else {
                        sb.AppendFormat("{0} 個檔案", data.Length);
                    }
                    notifyIcon.BalloonTipText = sb.ToString();
                } else if (dataObject.GetDataPresent(DataFormats.WaveAudio)) {
                    using (Stream data = dataObject.GetData(DataFormats.WaveAudio) as Stream)
                        notifyIcon.BalloonTipText = string.Format("聲音資料 ({0} 位元組)", data.Length);
                } else {
                    notifyIcon.BalloonTipTitle = "剪貼簿已被清理";
                    notifyIcon.BalloonTipText = "空白內容";
                }
                notifyIcon.ShowBalloonTip(5000);
            } catch (Exception) { }
        }

        private static void RecordHistory(IDataObject dataObject, ToolStripMenuItem parent, List<ToolStripMenuItem> list, HashSet<ToolStripMenuItem> recycle, EventHandler handleUseClick, EventHandler handleRemoveClick) {
            ToolStripMenuItem item;
            if (recycle.Count > 0) {
                item = recycle.First();
                recycle.Remove(item);
            } else {
                item = new ToolStripMenuItem();
                ToolStripMenuItem useItem = new ToolStripMenuItem("拿出放到剪貼簿 (&U)");
                useItem.Click += handleUseClick;
                useItem.Tag = item;
                ToolStripMenuItem removeButton = new ToolStripMenuItem("忘記 (&R)");
                removeButton.Click += handleRemoveClick;
                removeButton.Tag = item;
                item.DropDownItems.Add(useItem);
                item.DropDownItems.Add(removeButton);
                item.DropDownItems.Add(new ToolStripSeparator());
            }
            UpdateDisplay(dataObject, item);
            item.Tag = CloneDataObject(dataObject);
            item.Enabled = true;
            parent.DropDownItems.Insert(list.Count, item);
            list.Add(item);
        }

        private static DataObject CloneDataObject(IDataObject dataObject) {
            DataObject result = new DataObject();
            foreach (string format in dataObject.GetFormats(false))
                result.SetData(format, dataObject.GetData(format, false));
            return result;
        }
    }
}
