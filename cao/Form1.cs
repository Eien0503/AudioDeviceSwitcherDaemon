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

namespace cao
{
    public partial class Form1 : Form, IMMNotificationClient
    {
        private MMDeviceEnumerator enumerator; // 音訊裝置列舉器
        private MMDevice defaultDevice; // 預設音訊裝置
        private List<MMDevice> audioDevices; // 音訊裝置列表
        private int currentDeviceIndex; // 當前音訊裝置索引
        private const int HOTKEY_ID = 1; // 快捷鍵 ID
        private const int WM_HOTKEY = 0x0312; // 快捷鍵訊息
        private CoreAudioController audioController; // 音訊控制器

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk); // 註冊全域快捷鍵

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id); // 取消註冊全域快捷鍵

        public Form1()
        {
            InitializeComponent(); // 初始化元件
            InitializeAudioDevice(); // 初始化音訊裝置
            RegisterGlobalHotKey(); // 註冊全域快捷鍵
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
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_NONE, VK_MEDIA_PLAY_PAUSE); // 註冊全域快捷鍵
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                SwitchToNextAudioDevice(); // 切換到下一個音訊裝置
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
                UpdateDefaultAudioDevice(); // 更新預設音訊裝置
                //RestartApplication(); // 重新啟動程式
            }
        }

        public void OnDeviceAdded(string pwstrDeviceId)
        {
            //UpdateAudioDevices(); // 更新音訊裝置列表
            RestartApplication(); // 重新啟動程式
        }

        public void OnDeviceRemoved(string deviceId)
        {
            //UpdateAudioDevices(); // 更新音訊裝置列表
            RestartApplication(); // 重新啟動程式
        }
        private void RestartApplication()
        {
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
            RestartApplication(); // 重新啟動程式
            //UpdateAudioDevices(); // 更新音訊裝置列表
        }
        public void OnPropertyValueChanged(string pwstrDeviceId, NAudioPropertyKey key) { } // 音訊裝置屬性值改變事件

        private void label1_Click(object sender, EventArgs e)
        {

        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID); // 取消註冊全域快捷鍵
            base.OnFormClosing(e); // 呼叫基底類別的 OnFormClosing 方法
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