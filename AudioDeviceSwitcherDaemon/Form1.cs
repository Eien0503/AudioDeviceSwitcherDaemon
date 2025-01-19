using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using AudioSwitcher.AudioApi.CoreAudio;
using NAudioPropertyKey = NAudio.CoreAudioApi.PropertyKey;
using AudioSwitcherPropertyKey = AudioSwitcher.AudioApi.CoreAudio.PropertyKey;
using System.Reflection.Emit;
using Microsoft.Win32;
using System.IO;
using IWshRuntimeLibrary;

namespace AudioDeviceSwitcherDaemon
{
    public partial class Form1 : Form, IMMNotificationClient
    {
        private MMDeviceEnumerator enumerator; // 音訊裝置列舉器
        private MMDevice defaultDevice; // 預設音訊裝置
        private List<MMDevice> audioDevices; // 音訊裝置列表
        private int currentDeviceIndex; // 當前音訊裝置索引
        private const int HOTKEY_ID_PLAY_PAUSE = 1; // 播放暫停快捷鍵 ID
        private const int HOTKEY_ID_PAUSE_BREAK = 2; // Pause Break 快捷鍵 ID
        private const int HOTKEY_ID_CUSTOM = 3; // 自訂快捷鍵 ID
        private const int WM_HOTKEY = 0x0312; // 快捷鍵訊息
        private CoreAudioController audioController; // 音訊控制器
        private NotifyIcon notifyIcon; // 系統通知區域圖示
        private ContextMenuStrip contextMenu; // 右鍵選單
        private bool usePauseBreakKey; // 是否使用 Pause Break 鍵
        private bool useCustomKey; // 是否使用自訂快捷鍵


        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk); // 註冊全域快捷鍵

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id); // 取消註冊全域快捷鍵

        public Form1()
        {
            InitializeComponent(); // 初始化元件
            InitializeAudioDevice(); // 初始化音訊裝置
            LoadSettings(); // 載入設定
            RegisterGlobalHotKey(); // 註冊全域快捷鍵
            InitializeNotifyIcon(); // 初始化通知圖示
            SetStartup(); // 設定隨開機啟動
        }

        private void SetStartup()
        {
            try
            {
                string startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                string shortcutPath = System.IO.Path.Combine(startupFolderPath, "AudioDeviceSwitcherDaemon.lnk");

                if (!System.IO.File.Exists(shortcutPath))
                {
                    CreateShortcut(shortcutPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定隨開機啟動時發生錯誤: {ex.Message}");
            }
        }

        private void CreateShortcut(string shortcutPath)
        {
            var shell = new WshShell();
            var shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = Application.ExecutablePath;
            shortcut.WorkingDirectory = Application.StartupPath;
            shortcut.Save();
        }

        private void InitializeAudioDevice()
        {
            try
            {
                enumerator = new MMDeviceEnumerator(); // 建立音訊裝置列舉器
                enumerator.RegisterEndpointNotificationCallback(this); // 註冊端點通知回呼
                audioDevices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList(); // 列舉音訊端點
                currentDeviceIndex = audioDevices.FindIndex(d => d.ID == enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia).ID); // 找到預設音訊端點的索引
                audioController = new CoreAudioController(); // 建立音訊控制器
                UpdateDefaultAudioDevice(); // 更新預設音訊裝置
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化音訊裝置時發生錯誤: {ex.Message}");
            }
        }

        private void UpdateDefaultAudioDevice()
        {
            if (audioDevices.Count > 0 && currentDeviceIndex >= 0 && currentDeviceIndex < audioDevices.Count)
            {
                defaultDevice = audioDevices[currentDeviceIndex]; // 設定預設音訊裝置
                UpdateLabel(defaultDevice.FriendlyName); // 更新標籤
            }
            else
            {
                UpdateLabel("無可用的音訊裝置"); // 更新標籤
            }
        }

        private void UpdateLabel(string text)
        {
            if (label1.InvokeRequired)
            {
                label1.Invoke(new Action(() => label1.Text = text)); // 如果需要，使用 Invoke 更新標籤
            }
            else
            {
                label1.Text = text; // 直接更新標籤
            }
        }

        private void RegisterGlobalHotKey()
        {
            const uint MOD_NONE = 0x0000; // 無修飾鍵
            const uint VK_MEDIA_PLAY_PAUSE = 0xB3; // 媒體播放暫停鍵
            const uint VK_PAUSE = 0x13; // Pause Break 鍵

            if (usePauseBreakKey)
            {
                RegisterHotKey(this.Handle, HOTKEY_ID_PAUSE_BREAK, MOD_NONE, VK_PAUSE); // 註冊 Pause Break 快捷鍵
            }
            else if (useCustomKey && !string.IsNullOrEmpty(Properties.Settings.Default.CustomHotKey) &&
                     Enum.TryParse(Properties.Settings.Default.CustomHotKey, out Keys customKey))
            {
                RegisterHotKey(this.Handle, HOTKEY_ID_CUSTOM, MOD_NONE, (uint)customKey); // 註冊自訂快捷鍵
            }
            else
            {
                RegisterHotKey(this.Handle, HOTKEY_ID_PLAY_PAUSE, MOD_NONE, VK_MEDIA_PLAY_PAUSE); // 註冊播放暫停快捷鍵
            }
        }

        private void InitializeNotifyIcon()
        {
            notifyIcon = new NotifyIcon();
            contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("使用播放暫停鍵", null, (s, e) => SwitchHotKey(false));
            contextMenu.Items.Add("使用Pause Break鍵", null, (s, e) => SwitchHotKey(true));
            contextMenu.Items.Add("使用自訂快捷鍵", null, (s, e) => CustomizeHotKey());
            contextMenu.Items.Add("Exit", null, (s, e) => ExitApplication());

            notifyIcon.Icon = new Icon("icon.ico"); // 設定為程式的圖示
            notifyIcon.ContextMenuStrip = contextMenu;
            notifyIcon.Visible = true;
            notifyIcon.DoubleClick += (s, e) => this.Show();
            UpdateContextMenu(); // 更新右鍵選單
        }

        private void SwitchHotKey(bool usePauseBreak)
        {
            usePauseBreakKey = usePauseBreak;
            useCustomKey = false; // 重置自訂快捷鍵狀態
            UnregisterHotKey(this.Handle, HOTKEY_ID_PLAY_PAUSE); // 取消註冊播放暫停快捷鍵
            UnregisterHotKey(this.Handle, HOTKEY_ID_PAUSE_BREAK); // 取消註冊 Pause Break 快捷鍵
            UnregisterHotKey(this.Handle, HOTKEY_ID_CUSTOM); // 取消註冊自訂快捷鍵
            RegisterGlobalHotKey(); // 重新註冊快捷鍵
            UpdateContextMenu(); // 更新右鍵選單
            SaveSettings(); // 儲存設定
        }

        private void UpdateContextMenu()
        {
            if (contextMenu == null)
            {
                contextMenu = new ContextMenuStrip();
            }

            foreach (ToolStripMenuItem item in contextMenu.Items)
            {
                item.Checked = false;
            }

            if (contextMenu.Items.Count > 1 && usePauseBreakKey)
            {
                ((ToolStripMenuItem)contextMenu.Items[1]).Checked = true; // 勾選 Pause Break 鍵
            }
            else if (contextMenu.Items.Count > 0 && !usePauseBreakKey && !useCustomKey)
            {
                ((ToolStripMenuItem)contextMenu.Items[0]).Checked = true; // 勾選播放暫停鍵
            }

            // 更新 "使用自訂快捷鍵" 選項的文字
            if (contextMenu.Items.Count > 2)
            {
                var customHotKeyItem = (ToolStripMenuItem)contextMenu.Items[2];
                if (!string.IsNullOrEmpty(Properties.Settings.Default.CustomHotKey) &&
                    Enum.TryParse(Properties.Settings.Default.CustomHotKey, out Keys customKey))
                {
                    customHotKeyItem.Text = $"使用自訂快捷鍵 ({customKey})";
                    if (useCustomKey)
                    {
                        customHotKeyItem.Checked = true; // 勾選自訂快捷鍵
                    }
                }
                else
                {
                    customHotKeyItem.Text = "使用自訂快捷鍵 (null)";
                }
            }
            else
            {
                var customHotKeyItem = new ToolStripMenuItem("使用自訂快捷鍵 (null)", null, (s, e) => CustomizeHotKey());
                if (contextMenu.Items.Count >= 2)
                {
                    contextMenu.Items.Insert(2, customHotKeyItem);
                }
                else
                {
                    contextMenu.Items.Add(customHotKeyItem);
                }
            }
        }

        private void LoadSettings()
        {
            usePauseBreakKey = Properties.Settings.Default.UsePauseBreakKey;
            useCustomKey = Properties.Settings.Default.UseCustomKey;

            if (useCustomKey && Enum.TryParse(Properties.Settings.Default.CustomHotKey, out Keys customKey))
            {
                RegisterCustomHotKey(customKey);
            }
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.UsePauseBreakKey = usePauseBreakKey;
            Properties.Settings.Default.UseCustomKey = useCustomKey;
            Properties.Settings.Default.Save();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.Hide(); // 隱藏視窗
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                if ((m.WParam.ToInt32() == HOTKEY_ID_PLAY_PAUSE && !usePauseBreakKey && !useCustomKey) ||
                    (m.WParam.ToInt32() == HOTKEY_ID_PAUSE_BREAK && usePauseBreakKey) ||
                    (m.WParam.ToInt32() == HOTKEY_ID_CUSTOM && useCustomKey))
                {
                    SwitchToNextAudioDevice(); // 切換到下一個音訊裝置
                }
            }
            base.WndProc(ref m); // 呼叫基底類別的 WndProc 方法
        }

        private void SwitchToNextAudioDevice()
        {
            if (audioDevices.Count > 0)
            {
                currentDeviceIndex = (currentDeviceIndex + 1) % audioDevices.Count; // 更新當前音訊裝置索引
                SetDefaultAudioDevice(audioDevices[currentDeviceIndex].ID); // 設定預設音訊裝置
                UpdateDefaultAudioDevice(); // 更新預設音訊裝置
            }
        }

        private void SetDefaultAudioDevice(string deviceId)
        {
            var device = audioDevices.FirstOrDefault(d => d.ID == deviceId); // 找到對應的音訊裝置
            if (device != null)
            {
                SetDefaultAudioPlaybackDevice(deviceId); // 設定預設音訊播放裝置
            }
        }

        private void SetDefaultAudioPlaybackDevice(string deviceId)
        {
            try
            {
                var policyConfig = new PolicyConfigVistaClient(); // 建立 PolicyConfigVistaClient
                var policyConfigInterface = (IPolicyConfigVista)policyConfig; // 取得 IPolicyConfigVista 介面
                policyConfigInterface.SetDefaultEndpoint(deviceId, ERole.eMultimedia); // 設定預設端點
            }
            catch (InvalidCastException ex)
            {
                MessageBox.Show($"無法將 COM 物件轉換為 IPolicyConfigVista 介面: {ex.Message}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定預設音訊播放裝置時發生錯誤: {ex.Message}");
            }
        }

        public void OnDefaultDeviceChanged(DataFlow flow, Role role, string defaultDeviceId)
        {
            if (flow == DataFlow.Render && role == Role.Multimedia)
            {
                currentDeviceIndex = audioDevices.FindIndex(d => d.ID == defaultDeviceId); // 更新當前音訊裝置索引
                UpdateDefaultAudioDevice(); // 更新預設音訊裝置
                                            //RestartApplication(); // 重新啟動程式
            }
        }

        public void OnDeviceAdded(string pwstrDeviceId)
        {
            //UpdateAudioDevices(); // 更新音訊裝置列表
            //RestartApplication(); // 重新啟動程式
        }

        public void OnDeviceRemoved(string deviceId)
        {
            //UpdateAudioDevices(); // 更新音訊裝置列表
            //RestartApplication(); // 重新啟動程式
        }
        private void RestartApplication()
        {
            if (notifyIcon != null)
            {
                notifyIcon.Visible = false; // 隱藏通知圖示
                notifyIcon.Dispose(); // 釋放通知圖示資源
            }
            Application.ExitThread(); // 確保所有執行緒都被正確關閉
            Application.Restart(); // 重新啟動應用程式
            Environment.Exit(0); // 結束當前執行的應用程式
        }

        private void UpdateAudioDevices()
        {
            audioDevices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToList(); // 重新列舉音訊端點
            currentDeviceIndex = audioDevices.FindIndex(d => d.ID == enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia).ID); // 更新當前音訊裝置索引
            if (currentDeviceIndex == -1 && audioDevices.Count > 0)
            {
                currentDeviceIndex = 0; // 如果找不到預設音訊裝置，將索引設為 0
            }
            UpdateDefaultAudioDevice(); // 更新預設音訊裝置
        }

        public void OnDeviceStateChanged(string deviceId, DeviceState newState) // 當音訊裝置狀態改變事件
        {
            if (notifyIcon != null)
            {
                notifyIcon.Visible = false; // 隱藏通知圖示
                notifyIcon.Dispose(); // 釋放通知圖示資源
            }
            RestartApplication(); // 重新啟動程式
                                  //UpdateDefaultAudioDevice(); // 更新預設音訊裝置
                                  //UpdateAudioDevices(); // 更新音訊裝置列表
        }
        public void OnPropertyValueChanged(string pwstrDeviceId, NAudioPropertyKey key) { } // 音訊裝置屬性值改變事件

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void ExitApplication()
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID_PLAY_PAUSE); // 取消註冊播放暫停快捷鍵
            UnregisterHotKey(this.Handle, HOTKEY_ID_PAUSE_BREAK); // 取消註冊 Pause Break 快捷鍵

            if (notifyIcon != null)
            {
                notifyIcon.Visible = false; // 隱藏通知圖示
                notifyIcon.Dispose(); // 釋放通知圖示資源
            }

            Application.ExitThread(); // 確保所有執行緒都被正確關閉
            Application.Exit(); // 關閉應用程式
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            e.Cancel = true; // 取消關閉動作
            this.Hide(); // 隱藏視窗
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide(); // 隱藏視窗
            }
        }
        private int GetAudioDeviceCount()
        {
            return audioDevices.Count; // 回傳音訊裝置數量
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.Close(); // 關閉程式
                return true; // 表示已處理該按鍵
            }
            return base.ProcessCmdKey(ref msg, keyData); // 呼叫基底類別的 ProcessCmdKey 方法
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void CustomizeHotKey()
        {
            using (var form = new Form())
            {
                var label = new System.Windows.Forms.Label() { Left = 10, Top = 20, Text = "輸入新快捷鍵:" };
                var textBox = new TextBox() { Left = 10, Top = 50, Width = 200 };
                var buttonOk = new Button() { Text = "確定", Left = 10, Width = 100, Top = 80, DialogResult = DialogResult.OK };
                var buttonCancel = new Button() { Text = "取消", Left = 120, Width = 100, Top = 80, DialogResult = DialogResult.Cancel };
                var buttonClear = new Button() { Text = "取消自訂快捷鍵", Left = 230, Width = 100, Top = 80 };

                // 初始化 TextBox 的內容為當前的自訂快捷鍵
                if (!string.IsNullOrEmpty(Properties.Settings.Default.CustomHotKey) &&
                    Enum.TryParse(Properties.Settings.Default.CustomHotKey, out Keys customKey))
                {
                    textBox.Text = $"按鍵名稱: {customKey}";
                    textBox.Tag = customKey;
                }

                textBox.KeyDown += (sender, e) => TextBox_KeyDown(sender, e, textBox);
                buttonClear.Click += (sender, e) => ClearCustomHotKey(textBox);

                form.Text = "自訂快捷鍵";
                form.ClientSize = new Size(350, 120);
                form.Icon = new Icon("icon.ico"); // 設定視窗圖示
                form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel, buttonClear });
                form.AcceptButton = buttonOk;
                form.CancelButton = buttonCancel;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    if (string.IsNullOrEmpty(textBox.Text) || !Enum.TryParse(textBox.Tag?.ToString(), out Keys newKey))
                    {
                        if (!string.IsNullOrEmpty(textBox.Text))
                        {
                            MessageBox.Show("無效的快捷鍵");
                        }
                    }
                    else
                    {
                        RegisterCustomHotKey(newKey);
                    }
                }
            }
        }

        private void ClearCustomHotKey(TextBox textBox)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID_CUSTOM); // 取消註冊自訂快捷鍵

            Properties.Settings.Default.CustomHotKey = string.Empty;
            Properties.Settings.Default.UseCustomKey = false; // 重置自訂快捷鍵狀態
            Properties.Settings.Default.Save();

            textBox.Clear(); // 清除 TextBox 內容

            useCustomKey = false; // 重置自訂快捷鍵狀態

            // 更新右鍵選單，但不顯示其他快捷鍵的勾選
            foreach (ToolStripMenuItem item in contextMenu.Items)
            {
                item.Checked = false;
            }

            var customHotKeyItem = (ToolStripMenuItem)contextMenu.Items[2];
            customHotKeyItem.Text = "使用自訂快捷鍵 (null)";
            customHotKeyItem.Checked = false;
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e, TextBox textBox)
        {
            string keyName = e.KeyCode.ToString();
            textBox.Text = $"按鍵名稱: {keyName}";
            textBox.Tag = e.KeyCode; // 將鍵盤名稱存儲在 TextBox 的 Tag 屬性中
            e.SuppressKeyPress = true; // 防止 TextBox 顯示按鍵字符
        }
        private void RegisterCustomHotKey(Keys newKey)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID_PLAY_PAUSE); // 取消註冊播放暫停快捷鍵
            UnregisterHotKey(this.Handle, HOTKEY_ID_PAUSE_BREAK); // 取消註冊 Pause Break 快捷鍵
            UnregisterHotKey(this.Handle, HOTKEY_ID_CUSTOM); // 取消註冊自訂快捷鍵

            const uint MOD_NONE = 0x0000; // 無修飾鍵
            RegisterHotKey(this.Handle, HOTKEY_ID_CUSTOM, MOD_NONE, (uint)newKey); // 註冊新的快捷鍵

            Properties.Settings.Default.CustomHotKey = newKey.ToString();
            Properties.Settings.Default.Save();

            useCustomKey = true; // 設置自訂快捷鍵狀態
            usePauseBreakKey = false; // 重置 Pause Break 狀態

            UpdateContextMenu(); // 更新右鍵選單
            SaveSettings(); // 儲存設定
        }
    }

    [ComImport, Guid("294935CE-F637-4E7C-A41B-AB255460B862")]
    internal class PolicyConfigVistaClient
    {
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("568b9108-44bf-40b4-9006-86afe5b5a620")]
    internal interface IPolicyConfigVista
    {
        void GetMixFormat(); // 取得混合格式
        void GetDeviceFormat(); // 取得裝置格式
        void SetDeviceFormat(); // 設定裝置格式
        void GetProcessingPeriod(); // 取得處理週期
        void SetProcessingPeriod(); // 設定處理週期
        void GetShareMode(); // 取得共用模式
        void SetShareMode(); // 設定共用模式
        void GetPropertyValue(); // 取得屬性值
        void SetPropertyValue(); // 設定屬性值
        void SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string wszDeviceId, ERole eRole); // 設定預設端點
        void SetEndpointVisibility(); // 設定端點可見性
    }

    internal enum ERole
    {
        eConsole = 0, // 控制台
        eMultimedia = 1, // 多媒體
        eCommunications = 2, // 通訊
        ERole_enum_count = 3 // 列舉計數
    }
}