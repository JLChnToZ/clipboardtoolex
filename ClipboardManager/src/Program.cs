using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Mutex = System.Threading.Mutex;

namespace ClipboardManager {
    internal static class Program {
        private static NotifyIcon notifyIcon;
        private static Timer cleanupTimer;
        private static ClipboardNotifier clipboardNotifier;
        private static ToolStripMenuItem dataDisplay;
        private static ToolStripLabel imageDataDisplay;
        private static ToolStripMenuItem pinnedList, historyList;
        private static bool cleanUpEnabled = true;
        private static bool isRunning = false;
        private static readonly List<ToolStripMenuItem>
            pinnedItems = new List<ToolStripMenuItem>(),
            historyItems = new List<ToolStripMenuItem>();
        private static readonly HashSet<ToolStripMenuItem>
            unusedPinnedItems = new HashSet<ToolStripMenuItem>(),
            unusedHistoryItems = new HashSet<ToolStripMenuItem>();

        private static int MaxHistoryObjects {
            get { return Properties.Settings.Default.maxHistoryObjects; }
            set {
                Properties.Settings.Default.maxHistoryObjects = value;
                Properties.Settings.Default.Save();
            }
        }
        private static bool NotifyEnabled {
            get { return Properties.Settings.Default.notifyEnabled; }
            set {
                Properties.Settings.Default.notifyEnabled = value;
                Properties.Settings.Default.Save();
            }
        }

        [STAThread]
        private static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Mutex mutex = null;
            try {
                mutex = new Mutex(false, (Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(GuidAttribute), true)[0] as GuidAttribute).Value);
                Application.Idle += Run;
                Application.Run();
                isRunning = true;
            } catch (Exception ex) {
                if(!isRunning) Application.Run();
                ShowErrorMessage(ex);
            } finally {
                mutex?.Dispose();
            }
        }

        private static void Run(object sender, EventArgs e) {
            Application.Idle -= Run;

            clipboardNotifier = new ClipboardNotifier();
            clipboardNotifier.ClipboardChanged += HandleClipboard;

            cleanupTimer = new Timer {
                Interval = Properties.Settings.Default.autoCleanInterval
            };
            cleanUpEnabled = cleanupTimer.Interval > 0;
            cleanupTimer.Tick += HandleCleanUpClicked;

            notifyIcon = new NotifyIcon {
                Icon = SystemIcons.Application,
                Text = Application.ProductName,
            };
            notifyIcon.DoubleClick += HandleCleanUpClicked;

            ContextMenuStrip iconMenu = new ContextMenuStrip();
            notifyIcon.ContextMenuStrip = iconMenu;

            iconMenu.Items.Add(new ToolStripLabel(string.Format("{0} - 版本 {1}", Application.ProductName, Application.ProductVersion)));
            iconMenu.Items.Add(new ToolStripSeparator());

            iconMenu.Items.Add(new ToolStripLabel("剪貼簿現有資料"));
            dataDisplay = new ToolStripMenuItem {
                Visible = false
            };
            iconMenu.Items.Add(dataDisplay);

            imageDataDisplay = new ToolStripLabel {
                Visible = false,
                AutoSize = false,
                BackgroundImageLayout = ImageLayout.Zoom
            };
            iconMenu.Items.Add(imageDataDisplay);

            ToolStripMenuItem cleanItem = new ToolStripMenuItem("立即清理 (&C)");
            cleanItem.Click += HandleCleanUpClicked;
            iconMenu.Items.Add(cleanItem);

            ToolStripMenuItem pinItems = new ToolStripMenuItem("記住 (&P)");
            pinItems.Click += HandlePinClick;
            iconMenu.Items.Add(pinItems);

            iconMenu.Items.Add(new ToolStripSeparator());

            iconMenu.Items.Add(new ToolStripLabel("自動在…秒後清理 (設為 0 則不自動清理) (&A)"));
            NumericUpDown intervalItemBox = new NumericUpDown {
                Minimum = 0,
                Maximum = int.MaxValue,
                Value = cleanupTimer.Interval / 1000,
            };
            intervalItemBox.ValueChanged += HandleIntervalChange;
            iconMenu.Items.Add(new ToolStripControlHost(intervalItemBox));

            iconMenu.Items.Add(new ToolStripLabel("自動記住…次記錄 (&R)"));
            NumericUpDown historyCountBox = new NumericUpDown {
                Minimum = 0,
                Maximum = int.MaxValue,
                Value = MaxHistoryObjects,
            };
            historyCountBox.ValueChanged += HandleHistoryCountChange;
            iconMenu.Items.Add(new ToolStripControlHost(historyCountBox));

            iconMenu.Items.Add(new ToolStripSeparator());

            pinnedList = new ToolStripMenuItem("已手動記住 (&I)");
            iconMenu.Items.Add(pinnedList);

            pinnedList.DropDownItems.Add(new ToolStripSeparator());
            ToolStripMenuItem clearPinnedItem = new ToolStripMenuItem("清理 (&C)");
            clearPinnedItem.Click += HandleClearPinnedClick;
            pinnedList.DropDownItems.Add(clearPinnedItem);
            
            historyList = new ToolStripMenuItem("歷史記錄 (&H)") {
                Enabled = MaxHistoryObjects > 0
            };
            iconMenu.Items.Add(historyList);

            historyList.DropDownItems.Add(new ToolStripSeparator());
            ToolStripMenuItem clearHistoryItem = new ToolStripMenuItem("清理 (&C)");
            clearHistoryItem.Click += HandleClearHistoryClick;
            historyList.DropDownItems.Add(clearHistoryItem);

            iconMenu.Items.Add(new ToolStripSeparator());

            ToolStripMenuItem notifyItem = new ToolStripMenuItem("提示剪貼簿被更改 (&N)") {
                Checked = NotifyEnabled,
            };
            notifyItem.Click += HandleNotifyToggle;
            iconMenu.Items.Add(notifyItem);

            iconMenu.Items.Add(new ToolStripSeparator());

            ToolStripMenuItem runStartUpItem = new ToolStripMenuItem("開機時啟動 (&S)") {
                Checked = RunAtStartup
            };
            runStartUpItem.Click += RunAtStartupClick;
            iconMenu.Items.Add(runStartUpItem);

            ToolStripMenuItem exitItem = new ToolStripMenuItem("離開 (&E)");
            exitItem.Click += HandleExitClick;
            iconMenu.Items.Add(exitItem);

            notifyIcon.Visible = true;
            HandleClipboard(false, Clipboard.GetDataObject());
        }

        #region Event Handlers
        private static void HandleClipboard(bool notify, IDataObject dataObject, bool startTimer = true) {
            bool hasData = UpdateDisplay(dataObject, dataDisplay, imageDataDisplay);
            if(notify) NotifyData(dataObject);
            if(hasData) {
                if (startTimer && cleanUpEnabled) cleanupTimer.Start();
                if (historyList.HasDropDownItems) {
                    ClearHistory(historyList, historyItems, unusedHistoryItems);
                    if (MaxHistoryObjects > 0)
                        RecordHistory(dataObject, historyList, historyItems, unusedHistoryItems, HandleUseHistoryClick, HandleRemoveHistoryClick);
                }
            }
        }

        private static void HandleIntervalChange(object sender, EventArgs e) {
            NumericUpDown interval = sender as NumericUpDown;
            cleanupTimer.Stop();
            cleanUpEnabled = interval.Value > 0;
            Properties.Settings.Default.autoCleanInterval = (int)interval.Value * 1000;
            if (cleanUpEnabled) {
                cleanupTimer.Interval = Properties.Settings.Default.autoCleanInterval;
                HandleClipboard(false, Clipboard.GetDataObject());
            }
            Properties.Settings.Default.Save();
        }

        private static void HandleHistoryCountChange(object sender, EventArgs e) {
            MaxHistoryObjects = (int)(sender as NumericUpDown).Value;
            ClearHistory(historyList, historyItems, unusedHistoryItems);
            historyList.Enabled = MaxHistoryObjects > 0;
        }

        private static void HandleNotifyToggle(object sender, EventArgs e) {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            item.Checked = NotifyEnabled = !item.Checked;
        }

        private static void HandleClipboard(object sender, ClipboardChangedEventArgs e) {
            cleanupTimer.Stop();
            HandleClipboard(NotifyEnabled, e.dataObject);
        }

        private static void HandleCleanUpClicked(object sender, EventArgs e) {
            try {
                Clipboard.Clear();
            } catch (Exception ex) {
                ShowErrorMessage(ex);
            }
        }

        private static void HandleExitClick(object sender, EventArgs e) {
            Application.Exit();
        }

        private static void HandlePinClick(object sender, EventArgs e) {
            RecordHistory(Clipboard.GetDataObject(), pinnedList, pinnedItems, unusedPinnedItems, HandleUsePinnedClick, HandleRemovePinnedClick);
        }

        private static void HandleUsePinnedClick(object sender, EventArgs e) {
            ToolStripMenuItem item = (sender as ToolStripMenuItem).Tag as ToolStripMenuItem;
            IDataObject dataObject = item.Tag as IDataObject;
            try {
                Clipboard.SetDataObject(dataObject, true);
            } catch (Exception ex) {
                ShowErrorMessage(ex);
            }
        }

        private static void HandleRemovePinnedClick(object sender, EventArgs e) {
            RemoveHistory((sender as ToolStripMenuItem).Tag as ToolStripMenuItem, pinnedList, pinnedItems, unusedPinnedItems);
        }

        private static void HandleClearPinnedClick(object sender, EventArgs e) {
            ClearHistory(pinnedList, pinnedItems, unusedPinnedItems, true);
        }

        private static void HandleUseHistoryClick(object sender, EventArgs e) {
            ToolStripMenuItem item = (sender as ToolStripMenuItem).Tag as ToolStripMenuItem;
            IDataObject dataObject = item.Tag as IDataObject;
            RemoveHistory(item, historyList, historyItems, unusedHistoryItems);
            try {
                Clipboard.SetDataObject(dataObject, true);
            } catch (Exception ex) {
                ShowErrorMessage(ex);
            }
        }
        
        private static void HandleRemoveHistoryClick(object sender, EventArgs e) {
            RemoveHistory((sender as ToolStripMenuItem).Tag as ToolStripMenuItem, historyList, historyItems, unusedHistoryItems);
        }

        private static void HandleClearHistoryClick(object sender, EventArgs e) {
            ClearHistory(historyList, historyItems, unusedHistoryItems, true);
        }

        private static void HandleFileClick(object sender, EventArgs e) {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            try {
                if (item.Tag != null)
                    Process.Start(new ProcessStartInfo(item.Tag as string) {
                        UseShellExecute = true
                    });
            } catch(Exception ex) {
                ShowErrorMessage(ex);
            }
        }

        private static void RunAtStartupClick(object sender, EventArgs e) {
            try {
                ToolStripMenuItem item = sender as ToolStripMenuItem;
                item.Checked = RunAtStartup = !item.Checked;
            } catch(Exception ex) {
                ShowErrorMessage(ex);
            }
        }
        #endregion

        private static void RemoveHistory(ToolStripMenuItem item, ToolStripMenuItem parent, List<ToolStripMenuItem> list, HashSet<ToolStripMenuItem> recycle) {
            item.Tag = null;
            parent.DropDownItems.Remove(item);
            list.RemoveAt(0);
            recycle.Add(item);
        }

        private static void ClearHistory(ToolStripMenuItem parent, List<ToolStripMenuItem> list, HashSet<ToolStripMenuItem> recycle, bool cleanAll = false) {
            ToolStripItemCollection children = parent.DropDownItems;
            int cleanCount = cleanAll ? 0 : Math.Max(0, MaxHistoryObjects - 1);
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

        private static void NotifyData(IDataObject dataObject) {
            if (!NotifyEnabled) return;
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
            foreach(string format in dataObject.GetFormats(false))
                result.SetData(format, dataObject.GetData(format, false));
            return result;
        }

        private static void ShowErrorMessage(Exception ex) {
            MessageBox.Show(
                ex.Message,
                Application.ProductName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1
            );
        }

        private static bool RunAtStartup {
            get {
                using(RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false)) {
                    object value = key.GetValue(Application.ProductName);
                    return value is string && string.Equals(value as string, Application.ExecutablePath, StringComparison.Ordinal);
                }
            }
            set {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)) {
                    if (value)
                        key.SetValue(Application.ProductName, Application.ExecutablePath);
                    else
                        key.DeleteValue(Application.ProductName);
                }
            }
        }
    }
}
