using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

using ClipboardManager.Properties;

using TSItem = System.Windows.Forms.ToolStripItem;
using TSItems = System.Windows.Forms.ToolStripItemCollection;
using TSMenuItem = System.Windows.Forms.ToolStripMenuItem;
using TSLabel = System.Windows.Forms.ToolStripLabel;
using TSSeparator = System.Windows.Forms.ToolStripSeparator;

namespace ClipboardManager {
    internal class ClipboardApplication {
        private readonly Settings settings = Settings.Default;
        private readonly NotifyIcon notifyIcon = new NotifyIcon {
            Icon = Language.mainicon,
            Text = Application.ProductName,
            ContextMenuStrip = new ContextMenuStrip(),
            Visible = true,
        };
        private readonly Timer cleanupTimer = new Timer {
            Enabled = false,
        };
        private readonly ClipboardNotifier clipboardNotifier = new ClipboardNotifier();
        private readonly TSMenuItem dataDisplay = new TSMenuItem {
            Visible = false,
        };
        private readonly TSLabel imageDataDisplay = new TSLabel {
            Visible = false,
            AutoSize = false,
            BackgroundImageLayout = ImageLayout.Zoom,
        };
        private readonly HistoryMenu
            pinned,
            history;

        private const string RUN_KEY = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
        private static bool RunAtStartup {
            get {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RUN_KEY, false))
                    return key.GetValue(Application.ProductName) is string value &&
                        string.Equals(value, Application.ExecutablePath, StringComparison.Ordinal);
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

            if (settings.autoCleanInterval > 0)
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

            TSItems children = notifyIcon.ContextMenuStrip.Items;
            children.AddRange(new TSItem[] {
                new TSLabel(string.Format(Language.Caption, Application.ProductName, Application.ProductVersion)),

                new TSSeparator(),

                new TSLabel(Language.ExistsDataCaption),
                dataDisplay,
                imageDataDisplay,
                new TSMenuItem(Language.ClearNow, null, HandleCleanUpClicked),
                new TSMenuItem(Language.Pin, null, HandlePinClick),

                new TSSeparator(),
            });

            pinned = new HistoryMenu(children, Language.PinnedMenu);

            children.Add(new TSSeparator());

            history = new HistoryMenu(children, Language.HistoryMenu) {
                removeOnUse = true
            };

            children.AddRange(new TSItem[] {
                new TSSeparator(),

                new TSLabel(Language.AutoCleanAt),
                new ToolStripControlHost(intervalItemBox),
                new TSLabel(Language.EnableHistory),
                new ToolStripControlHost(historyCountBox),
                new TSMenuItem(Language.NotifyChanges, null, HandleNotifyToggle) {
                    Checked = settings.notifyEnabled,
                },
                new TSMenuItem(Language.RunAtStartup, null, RunAtStartupClick) {
                    Checked = RunAtStartup
                },

                new TSSeparator(),

                new TSMenuItem(Language.Exit, null, HandleExitClick),
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
                history.ClearHistory(Math.Max(0, settings.maxHistoryObjects - 1));
                if (settings.maxHistoryObjects > 0)
                    history.RecordHistory(dataObject);
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
            history.ClearHistory(settings.maxHistoryObjects = (int)(sender as NumericUpDown).Value);
            settings.Save();
        }

        private void HandleNotifyToggle(object sender, EventArgs e) {
            TSMenuItem item = sender as TSMenuItem;
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
            pinned.RecordHistory();
        }

        private static void HandleFileClick(object sender, EventArgs e) {
            TSMenuItem item = sender as TSMenuItem;
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
                TSMenuItem item = sender as TSMenuItem;
                item.Checked = RunAtStartup = !item.Checked;
            } catch (Exception ex) {
                UIHelper.ShowErrorMessage(ex);
            }
        }

        private void HandleApplicationExit(object sender, EventArgs e) {
            notifyIcon.Visible = false;
        }
        #endregion

        internal static bool UpdateDisplay(IDataObject dataObject, TSMenuItem dataDisplay, TSLabel imageDataDisplay = null) {
            bool hasData = false;
            try {
                hasData = true;
                dataDisplay.Visible = true;
                dataDisplay.Enabled = false;
                dataDisplay.ToolTipText = string.Empty;
                dataDisplay.Text = string.Empty;
                if (dataDisplay.HasDropDownItems) {
                    if (imageDataDisplay == null) {
                        HashSet<TSItem> pendingRemoveItems = new HashSet<TSItem>();
                        foreach (TSItem item in dataDisplay.DropDownItems)
                            if (item is TSLabel) {
                                imageDataDisplay = item as TSLabel;
                                break;
                            } else if (item.Tag is string)
                                pendingRemoveItems.Add(item);
                        foreach (TSItem item in pendingRemoveItems) {
                            if (item.Image != null) {
                                item.Image.Dispose();
                                item.Image = null;
                            }
                            dataDisplay.DropDownItems.Remove(item);
                        }
                    } else {
                        foreach (TSItem item in dataDisplay.DropDownItems)
                            if (item.Image != null) {
                                item.Image.Dispose();
                                item.Image = null;
                            }
                        dataDisplay.DropDownItems.Clear();
                    }
                }
                if (imageDataDisplay != null) {
                    imageDataDisplay.Visible = false;
                    if (imageDataDisplay.BackgroundImage != null) {
                        imageDataDisplay.BackgroundImage.Dispose();
                        imageDataDisplay.BackgroundImage = null;
                    }
                }
                if (dataObject.GetDataPresent(DataFormats.UnicodeText, true)) {
                    string data = dataObject.GetData(DataFormats.UnicodeText, true) as string;
                    string formattedData = data.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
                    if (formattedData.Length > 30)
                        dataDisplay.Text = string.Format(Language.TextData, formattedData.Substring(0, 30));
                    else
                        dataDisplay.Text = formattedData;
                    if (data.Length > 512)
                        dataDisplay.ToolTipText = string.Format(Language.TextData, data.Substring(0, 512));
                    else
                        dataDisplay.ToolTipText = data;
                } else if (dataObject.GetDataPresent(DataFormats.Bitmap)) {
                    Image data = dataObject.GetData(DataFormats.Bitmap) as Image;
                    string imgMeta = string.Format(Language.ImageData, data.Width, data.Height);
                    if (imageDataDisplay == null) {
                        imageDataDisplay = new TSLabel {
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
                    dataDisplay.Text = string.Format(Language.FileData, data.Length);
                    dataDisplay.Enabled = true;
                    foreach (string fileName in data) {
                        Bitmap icon = null;
                        try {
                            using (Icon rawIcon = Native.GetSmallIcon(fileName))
                                icon = rawIcon.ToBitmap();
                        } catch (Exception) { }
                        TSMenuItem item = new TSMenuItem(Path.GetFileName(fileName)) {
                            Tag = fileName,
                            Image = icon,
                        };
                        item.Click += HandleFileClick;
                        dataDisplay.DropDownItems.Add(item);
                    }
                } else if (dataObject.GetDataPresent(DataFormats.WaveAudio)) {
                    using (Stream data = dataObject.GetData(DataFormats.WaveAudio) as Stream)
                        dataDisplay.Text = string.Format(Language.AudioData, data.Length);
                } else {
                    string[] formats = dataObject.GetFormats(false);
                    if (formats.Length > 0) {
                        dataDisplay.Text = string.Join(", ", formats);
                    } else {
                        dataDisplay.Visible = false;
                        hasData = false;
                    }
                }
            } catch (Exception) { }
            return hasData;
        }

        private void NotifyData(IDataObject dataObject) {
            if (!settings.notifyEnabled) return;
            try {
                notifyIcon.BalloonTipTitle = Language.OnChangeTitle;
                if (dataObject.GetDataPresent(DataFormats.UnicodeText, true)) {
                    string data = dataObject.GetData(DataFormats.UnicodeText, true) as string;
                    if (data.Length > 64)
                        notifyIcon.BalloonTipText = string.Format(Language.OnChangeTextData, data.Substring(0, 64), data.Length - 64);
                    else
                        notifyIcon.BalloonTipText = data;
                } else if (dataObject.GetDataPresent(DataFormats.Bitmap)) {
                    Image data = dataObject.GetData(DataFormats.Bitmap) as Image;
                    notifyIcon.BalloonTipText = string.Format(Language.OnChangeImageData, data.Width, data.Height);
                } else if (dataObject.GetDataPresent(DataFormats.FileDrop)) {
                    string[] data = dataObject.GetData(DataFormats.FileDrop) as string[];
                    StringBuilder sb = new StringBuilder();
                    if (data.Length < 3 && data.Length > 0) {
                        foreach (string fileName in data) {
                            if (sb.Length > 0) sb.Append(", ");
                            sb.Append(Path.GetFileName(fileName));
                        }
                    } else {
                        sb.AppendFormat(Language.OnChangeFileData, data.Length);
                    }
                    notifyIcon.BalloonTipText = sb.ToString();
                } else if (dataObject.GetDataPresent(DataFormats.WaveAudio)) {
                    using (Stream data = dataObject.GetData(DataFormats.WaveAudio) as Stream)
                        notifyIcon.BalloonTipText = string.Format(Language.OnChangeAudioData, data.Length);
                } else {
                    string[] formats = dataObject.GetFormats(false);
                    if (formats.Length > 0) {
                        notifyIcon.BalloonTipText = string.Join(", ", formats);
                    } else {
                        notifyIcon.BalloonTipTitle = Language.OnClearTitle;
                        notifyIcon.BalloonTipText = Language.OnClearDescription;
                    }
                }
                notifyIcon.ShowBalloonTip(5000);
            } catch (Exception) { }
        }
    }
}
