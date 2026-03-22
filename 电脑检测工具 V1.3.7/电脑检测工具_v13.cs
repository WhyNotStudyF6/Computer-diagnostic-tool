using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using Microsoft.Win32;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace PcCheckTool
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            var startupMode = ParseStartupMode(Environment.GetCommandLineArgs().Skip(1));
            Application.ThreadException += (_, eventArgs) => ShowFatalError(eventArgs == null ? null : eventArgs.Exception);
            AppDomain.CurrentDomain.UnhandledException += (_, eventArgs) =>
            {
                var exception = eventArgs == null ? null : eventArgs.ExceptionObject as Exception;
                ShowFatalError(exception ?? new Exception("程序发生未处理异常。"));
            };
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(startupMode));
        }

        private static StartupMode ParseStartupMode(IEnumerable<string> args)
        {
            foreach (var rawArg in args ?? Enumerable.Empty<string>())
            {
                var matchedMode = MatchStartupModeToken((rawArg ?? string.Empty).Trim().TrimStart('-', '/'));
                if (matchedMode != StartupMode.None)
                {
                    return matchedMode;
                }
            }

            var executableName = Path.GetFileNameWithoutExtension(AppDomain.CurrentDomain.FriendlyName ?? string.Empty);
            var fileNameMode = MatchStartupModeToken(executableName);
            if (fileNameMode != StartupMode.None)
            {
                return fileNameMode;
            }

            return StartupMode.None;
        }

        private static StartupMode MatchStartupModeToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return StartupMode.None;
            }

            if (token.IndexOf("zen", StringComparison.OrdinalIgnoreCase) >= 0
                || token.IndexOf("禅模式", StringComparison.OrdinalIgnoreCase) >= 0
                || token.IndexOf("禅", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return StartupMode.Zen;
            }

            if (token.IndexOf("speedrun", StringComparison.OrdinalIgnoreCase) >= 0
                || token.IndexOf("quick", StringComparison.OrdinalIgnoreCase) >= 0
                || token.IndexOf("速通版", StringComparison.OrdinalIgnoreCase) >= 0
                || token.IndexOf("速通", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return StartupMode.Speedrun;
            }

            return StartupMode.None;
        }

        private static void ShowFatalError(Exception exception)
        {
            try
            {
                var message = exception == null
                    ? "程序发生未处理异常。"
                    : string.Format("程序出现未处理错误：{0}\r\n\r\n{1}", exception.Message, exception);
                MessageBox.Show(message, "电脑检测工具", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch
            {
            }
        }
    }

    internal sealed class MainForm : Form
    {
        private const string AppTitle = "电脑检测工具";
        private const string AppVersion = "V1.3.7";
        private const string DefaultSoftwareUpdatedAt = "2026-03-22";
        private const string LinkEditPassword = "bkandawh";
        private const string DefaultWindowsActivationCopyTemplate = "slmgr /ipk {key}";
        private const string DefaultOfficeCommunicationTestUrl = "http://10.229.159.132:8000";
        private const string DefaultOfficeCommunicationTestPayload = "Office 通信测试\r\n只发送模拟文本，不包含激活数据。";
        private const string DefaultOfficeSequenceMenuText = "发送按键串";
        private const string DefaultOfficePostActionCondition = "关闭激活弹窗后";
        private const string DefaultOfficePostActionSequence =
            "按键:{Alt}\r\n" +
            "停顿:250\r\n" +
            "按键:D\r\n" +
            "停顿:400\r\n" +
            "按键:Y\r\n" +
            "停顿:200\r\n" +
            "按键:4\r\n" +
            "停顿:500\r\n" +
            "按键:{Tab}\r\n" +
            "停顿:250\r\n" +
            "按键:{Tab}\r\n" +
            "停顿:250\r\n" +
            "按键:{Tab}\r\n" +
            "停顿:500\r\n" +
            "按键:^v\r\n" +
            "停顿:500\r\n" +
            "按键:{Tab}\r\n" +
            "停顿:500\r\n" +
            "按键:{Enter}";
        private const string DefaultOfficePostAhkScript =
            "; 可用占位符: {key} {condition} {appdir}\r\n" +
            "; 下面只是示例, 你可以自行替换\r\n" +
            "SendInput, {Text}{key}";
        private const string OfficeKeySourcePath = "3.txt";
        private const string OfficeKeyTargetPath = "【重要】プロダクトキー情報とお問い合わせ先.txt";
        private const string OfficeKeyLabel = "Officeのプロダクトキー：";
        private const string OfficeKeySeparator = "＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝";
        private static readonly Regex OfficeProductKeyRegex = new Regex(@"\b[A-Z0-9]{5}(?:-[A-Z0-9]{5}){4}\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex OfficeKeyTitleRegex = new Regex(@"^[^\r\n]*Office[^\r\n]*プロダクトキー[^\r\n]*$", RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.CultureInvariant);
        private static readonly Regex OfficeKeySeparatorRegex = new Regex(@"^\s*[=＝\-－_＿~～\*＊#＃一—─━]{4,}\s*$", RegexOptions.Multiline | RegexOptions.CultureInvariant);
        private readonly AppSettings _settings = AppSettings.Load();
        private readonly List<Button> _trackedButtons = new List<Button>();
        private readonly Dictionary<float, Font> _actionButtonFontCache = new Dictionary<float, Font>();
        private Label _titleLabel;
        private Label _subLabel;
        private Panel _infoPanel;
        private Panel _batteryPanel;
        private Panel _statusPanel;
        private Panel _buttonsPanel;
        private TextBox _osLabel;
        private TextBox _modelLabel;
        private TextBox _cpuLabel;
        private TextBox _memoryLabel;
        private TextBox _diskLabel;
        private TextBox _batteryHealthLabel;
        private TextBox _batteryWearLabel;
        private TextBox _batteryCapacityLabel;
        private TextBox _batteryStatusLabel;
        private TextBox _wifiLabel;
        private TextBox _bluetoothLabel;
        private TextBox _panasonicLabel;
        private Button _wifiActionButton;
        private Button _windowsActionButton;
        private Button _configActionButton;
        private Button _batteryActionButton;
        private Button _cameraActionButton;
        private Button _keyboardActionButton;
        private Button _officeActionButton;
        private Button _officeOperationDropButton;
        private Button _clearActionButton;
        private Button _powerActionButton;
        private Button _settingsButton;
        private Button _zenModeButton;
        private Button _speedrunModeButton;
        private Button _exitButton;
        private CheckBox _topMostButton;
        private CheckBox _themeButton;
        private TextBox _noteTextBox;
        private Button _speakerTestButton;
        private Form _zenModeWindow;
        private Form _speedrunModeWindow;
        private NotifyIcon _modeNotifyIcon;
        private Icon _configuredAppIcon;
        private ContextMenuStrip _modeContextMenu;
        private ContextMenuStrip _officeOperationMenu;
        private ToolStripMenuItem _officeSequenceMenuItem;
        private ToolStripSeparator _officeSequenceMenuSeparator;
        private ContextMenuStrip _windowsActivationMenu;
        private ContextMenuStrip _wifiSelectionMenu;
        private ContextMenuStrip _suppressedContextMenu;
        private System.Windows.Forms.Timer _wifiStatusTimer;
        private System.Windows.Forms.Timer _officeActivationTimer;
        private readonly ToolTip _toolTip = new ToolTip();
        private string _computerManufacturer = string.Empty;
        private string _computerModel = string.Empty;
        private bool _completionShown;
        private double _batteryHealthPercent = -1d;
        private double _batteryFullChargeCapacityMWh;
        private double _batteryRemainingPercent = -1d;
        private bool _batteryPresent;
        private int _officeDesktopWriteCount;
        private int _officeCopyCount;
        private int _officeCutCount;
        private int _officeDeleteCount;
        private bool _officeActivationCheckInProgress;
        private string _lastOfficeActivationStatus = string.Empty;
        private string _lastWindowsActivationStatus = string.Empty;
        private string _lastWindowsTimeSyncStatus = string.Empty;
        private readonly StartupMode _startupMode;
        private bool _startupModeOpened;
        private bool _applicationExiting;
        private DisplayMode _currentDisplayMode = DisplayMode.Normal;
        private Rectangle _normalModeBounds = Rectangle.Empty;
        private bool _normalModeWasMaximized;

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        private static extern uint SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);

        [DllImport("wlanapi.dll")]
        private static extern uint WlanOpenHandle(
            uint dwClientVersion,
            IntPtr pReserved,
            out uint pdwNegotiatedVersion,
            out IntPtr phClientHandle);

        [DllImport("wlanapi.dll")]
        private static extern uint WlanCloseHandle(IntPtr hClientHandle, IntPtr pReserved);

        [DllImport("wlanapi.dll")]
        private static extern uint WlanQueryInterface(
            IntPtr hClientHandle,
            ref Guid pInterfaceGuid,
            WlanIntfOpcode opCode,
            IntPtr pReserved,
            out int pdwDataSize,
            out IntPtr ppData,
            IntPtr pWlanOpcodeValueType);

        [DllImport("wlanapi.dll")]
        private static extern uint WlanSetInterface(
            IntPtr hClientHandle,
            ref Guid pInterfaceGuid,
            WlanIntfOpcode opCode,
            int dwDataSize,
            ref WlanPhyRadioState pData,
            IntPtr pReserved);

        [DllImport("wlanapi.dll")]
        private static extern void WlanFreeMemory(IntPtr pMemory);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            uint Msg,
            IntPtr wParam,
            string lParam,
            uint fuFlags,
            uint uTimeout,
            out IntPtr lpdwResult);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const uint ShrbNoconfirmation = 0x00000001;
        private const uint ShrbNoprogressui = 0x00000002;
        private const uint ShrbNosound = 0x00000004;
        private const uint WmSettingchange = 0x001A;
        private const uint SmtoAbortifhung = 0x0002;
        private const uint KeyeventfKeyup = 0x0002;

        public MainForm(StartupMode startupMode)
        {
            _startupMode = startupMode;
            Text = GetConfiguredAppTitle();
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            ClientSize = new Size(600, 800);
            MinimumSize = new Size(300, 400);
            MaximumSize = Size.Empty;
            BackColor = Color.FromArgb(241, 245, 249);
            Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            ResizeRedraw = true;
            AutoScroll = true;
            if (_startupMode != StartupMode.None)
            {
                Opacity = 0d;
                ShowInTaskbar = false;
            }
            ApplyRememberedWindowBounds();

            if (string.IsNullOrWhiteSpace(_settings.OfficeKeySourcePath))
            {
                _settings.OfficeKeySourcePath = OfficeKeySourcePath;
            }

            if (string.IsNullOrWhiteSpace(_settings.OfficeKeyTargetPath))
            {
                _settings.OfficeKeyTargetPath = GetDefaultOfficeTargetPath();
            }

            if (_settings.OfficeKeyTargetPaths.Count == 0)
            {
                _settings.OfficeKeyTargetPaths.Add(_settings.OfficeKeyTargetPath);
            }

            if (string.IsNullOrWhiteSpace(_settings.KeyboardCheckPath))
            {
                _settings.KeyboardCheckPath = @"Keyboard Test Utility.exe";
            }

            if (string.IsNullOrWhiteSpace(_settings.SpeakerTestPath))
            {
                _settings.SpeakerTestPath = @"资源文件\【俺妹】俺の妹がこんなに可愛いわけがないop「irony」.mp4";
            }

            _settings.OfficeActivationCount = 0;

            _titleLabel = new Label
            {
                Text = GetConfiguredAppTitle(),
                Font = new Font("Microsoft YaHei UI", 20F, FontStyle.Bold, GraphicsUnit.Point),
                ForeColor = Color.FromArgb(15, 23, 42),
                Location = new Point(24, 18),
                Size = new Size(360, 42)
            };

            _subLabel = new Label
            {
                Text = "交机前快速核验系统、硬件与常用功能。",
                ForeColor = Color.FromArgb(71, 85, 105),
                Location = new Point(27, 62),
                Size = new Size(420, 24)
            };

            Controls.Add(_titleLabel);
            Controls.Add(_subLabel);

            _settingsButton = new Button
            {
                Text = "⚙",
                Location = new Point(662, 26),
                Size = new Size(40, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 23, 42),
                Font = new Font("Segoe UI Symbol", 11F, FontStyle.Regular, GraphicsUnit.Point)
            };
            _settingsButton.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            _settingsButton.Click += (_, __) => ShowSettingsDialog();
            _settingsButton.MouseUp += SettingsButtonMouseUp;
            Controls.Add(_settingsButton);
            _toolTip.SetToolTip(_settingsButton, "左击打开设置；右击打开系统个性化");

            _themeButton = new CheckBox
            {
                Text = "切换为暗主题",
                Appearance = Appearance.Button,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(540, 26),
                Size = new Size(116, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 23, 42),
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point),
                Checked = string.Equals(_settings.ThemeMode, "dark", StringComparison.OrdinalIgnoreCase)
            };
            _themeButton.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            _themeButton.CheckedChanged += (_, __) => ToggleTheme();
            _themeButton.MouseUp += ThemeButtonMouseUp;
            Controls.Add(_themeButton);
            _toolTip.SetToolTip(_themeButton, "左击切换明亮/黑暗主题，并同步系统与 Office 主题；右击打开系统个性化");

            _zenModeButton = CreateHeaderButton("禅模式", (_, __) => ShowZenModeWindow());
            _zenModeButton.Size = new Size(108, 30);
            Controls.Add(_zenModeButton);
            _toolTip.SetToolTip(_zenModeButton, "打开 300x400 的禅模式小窗口，只保留 1-9、扬声器测试、主题、返回和退出");

            _speedrunModeButton = CreateHeaderButton("速通版", (_, __) => ShowSpeedrunModeWindow());
            _speedrunModeButton.Size = new Size(108, 30);
            Controls.Add(_speedrunModeButton);
            _toolTip.SetToolTip(_speedrunModeButton, "打开 300x400 的速通版小窗口，只保留 Windows、配置、电池、Office、主题、返回和退出");

            _topMostButton = new CheckBox
            {
                Text = "⇧ 未置顶",
                Appearance = Appearance.Button,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(706, 26),
                Size = new Size(124, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 23, 42),
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                Checked = _settings.TopMost
            };
            _topMostButton.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            _topMostButton.CheckedChanged += (_, __) =>
            {
                TopMost = _topMostButton.Checked;
                _settings.TopMost = _topMostButton.Checked;
                _settings.Save();
                ApplyTopMostButtonVisualState();
                SyncModeWindowsTopMost();
            };
            ApplyTopMostButtonVisualState();
            Controls.Add(_topMostButton);
            _toolTip.SetToolTip(_topMostButton, "让本工具保持在最前面显示");
            TopMost = _settings.TopMost;

            _exitButton = null;

            _infoPanel = CreateCardPanel(new Point(24, 96), new Size(520, 340), "配置信息", "开机即显示当前系统和硬件概览");
            _osLabel = AddInfoRow(_infoPanel, "Windows", 60);
            _modelLabel = AddInfoRow(_infoPanel, "机型", 92);
            _cpuLabel = AddInfoRow(_infoPanel, "CPU", 124);
            _memoryLabel = AddInfoRow(_infoPanel, "内存", 156, 62);
            _diskLabel = AddInfoRow(_infoPanel, "硬盘", 226, 90);
            Controls.Add(_infoPanel);

            _batteryPanel = CreateCardPanel(new Point(564, 96), new Size(270, 340), "电池状态", "不点按钮也能直接看见");
            _batteryHealthLabel = AddInfoRow(_batteryPanel, "健康度", 60);
            _batteryWearLabel = AddInfoRow(_batteryPanel, "损耗", 92);
            _batteryCapacityLabel = AddInfoRow(_batteryPanel, "容量", 124);
            _batteryStatusLabel = AddInfoRow(_batteryPanel, "状态", 156);
            Controls.Add(_batteryPanel);

            _statusPanel = CreateCardPanel(new Point(24, 448), new Size(810, 142), "附加状态", "显示已保存 WiFi、蓝牙与机型相关驱动状态");
            _wifiLabel = AddInfoRow(_statusPanel, "WiFi", 64);
            _bluetoothLabel = AddInfoRow(_statusPanel, "蓝牙", 90);
            _panasonicLabel = AddInfoRow(_statusPanel, "松下圆环", 116);
            Controls.Add(_statusPanel);

            _buttonsPanel = CreateCardPanel(new Point(24, 608), new Size(810, 272), "快捷检测", "按钮直接执行操作，减少打断");
            _wifiActionButton = CreateActionButton("1. 联网", "按设置直接联网", new Point(22, 48), (sender, __) => ConnectWifiFromSettings(sender as Button), true);
            _buttonsPanel.Controls.Add(_wifiActionButton);
            _windowsActionButton = CreateActionButton("2. Windows 激活", "激活页 + 校准时间", new Point(282, 48), (_, __) => OpenWindowsActivation(), true);
            _buttonsPanel.Controls.Add(_windowsActionButton);
            _toolTip.SetToolTip(_windowsActionButton, "当前系统激活状态会自动检测。左击：打开激活页并校准时间；右击：显示管理员 CMD 激活菜单。");
            _configActionButton = CreateActionButton("3. 配置检查", "先开设置的系统-关于,再开磁盘管理", new Point(542, 48), (_, __) => OpenConfigurationCheck(), true);
            _buttonsPanel.Controls.Add(_configActionButton);
            _toolTip.SetToolTip(_configActionButton, "先打开设置的系统-关于，再打开磁盘管理。");
            _batteryActionButton = CreateActionButton("4. 电池检查", "打开 BatteryInfoView", new Point(22, 126), (_, __) => OpenBatteryCheck(), true);
            _buttonsPanel.Controls.Add(_batteryActionButton);
            _cameraActionButton = CreateActionButton("5. 摄像头检查", "打开相机看实时画面", new Point(282, 126), (_, __) => OpenCameraCheck(), true);
            _buttonsPanel.Controls.Add(_cameraActionButton);
            _keyboardActionButton = CreateActionButton("6. 键盘检查", "启动键盘测试工具", new Point(542, 126), (_, __) => OpenKeyboardCheck(), true);
            _buttonsPanel.Controls.Add(_keyboardActionButton);
            _officeActionButton = CreateActionButton("7. Office激活", "余0码 复制0次", new Point(22, 204), (sender, __) => OpenOfficeActivation(sender as Button), true);
            _buttonsPanel.Controls.Add(_officeActionButton);
            InitializeOfficeOperationMenu();
            InitializeWindowsActivationMenu();
            _officeOperationDropButton = new Button
            {
                Text = "›",
                Location = new Point(242, 212),
                Size = new Size(18, 18),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point),
                TabStop = false
            };
            _officeOperationDropButton.FlatAppearance.BorderSize = 1;
            _officeOperationDropButton.Click += (_, __) => ShowOfficeOperationMenu();
            _buttonsPanel.Controls.Add(_officeOperationDropButton);
            _officeOperationDropButton.BringToFront();
            _officeOperationDropButton.Visible = false;
            _officeOperationDropButton.Enabled = false;
            UpdateOfficeOperationDropButtonPresentation();
            _clearActionButton = CreateActionButton("8. 记录擦除", "清理最近/浏览器/回收站", new Point(282, 204), (sender, __) => ClearHistoryAndRecycleBin(sender as Button), true);
            _buttonsPanel.Controls.Add(_clearActionButton);
            _powerActionButton = CreateActionButton("9. 重启或关机", "执行前二次确认", new Point(542, 204), (_, __) => ShowPowerDialog(), false);
            _buttonsPanel.Controls.Add(_powerActionButton);
            Controls.Add(_buttonsPanel);

            _noteTextBox = new TextBox
            {
                Text = "颜色说明：灰色=未操作，橙色=已点击待确认，绿色对勾=通过，红色=异常/不合格；右击联网/Office可选项，其他按钮可手动标绿。\r\n当 1~8 全部变绿后，会进入自动关机倒计时界面，可取消。",
                Location = new Point(28, 892),
                Size = new Size(650, 42),
                ForeColor = Color.FromArgb(71, 85, 105),
                BackColor = BackColor,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                ShortcutsEnabled = true,
                Multiline = true,
                TabStop = false
            };

            Controls.Add(_noteTextBox);

            _speakerTestButton = new Button
            {
                Text = "扬声器测试",
                Location = new Point(706, 900),
                Size = new Size(108, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(219, 234, 254),
                ForeColor = Color.FromArgb(29, 78, 216),
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
            };
            _speakerTestButton.FlatAppearance.BorderColor = Color.FromArgb(147, 197, 253);
            _speakerTestButton.Click += (_, __) => OpenSpeakerTest();
            Controls.Add(_speakerTestButton);
            _toolTip.SetToolTip(_speakerTestButton, "播放扬声器测试视频，可在设置里修改路径");

            _modeNotifyIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = GetConfiguredAppTitle(),
                Visible = false
            };
            _modeNotifyIcon.DoubleClick += (_, __) => RestoreMainFromTray();
            InitializeModeContextMenu();
            AttachModeContextMenu(this);
            _modeNotifyIcon.ContextMenuStrip = _modeContextMenu;
            ApplyConfiguredAppIcon();

            Shown += (_, __) =>
            {
                ApplyRememberedWindowBounds();
                PerformResponsiveLayout();
                RefreshDashboard();
                BeginInvoke((MethodInvoker)TryOpenStartupMode);
            };

            _wifiStatusTimer = new System.Windows.Forms.Timer();
            _wifiStatusTimer.Interval = 4000;
            _wifiStatusTimer.Tick += (_, __) => SyncWifiActionButtonState();
            _wifiStatusTimer.Start();
            _officeActivationTimer = new System.Windows.Forms.Timer();
            _officeActivationTimer.Interval = 7000;
            _officeActivationTimer.Tick += (_, __) => SyncOfficeActivationButtonStateAsync();
            _officeActivationTimer.Start();
            FormClosed += (_, __) =>
            {
                _wifiStatusTimer.Stop();
                if (_officeActivationTimer != null)
                {
                    _officeActivationTimer.Stop();
                    _officeActivationTimer.Dispose();
                    _officeActivationTimer = null;
                }
                if (_officeOperationMenu != null)
                {
                    _officeOperationMenu.Dispose();
                    _officeOperationMenu = null;
                }
                if (_windowsActivationMenu != null)
                {
                    _windowsActivationMenu.Dispose();
                    _windowsActivationMenu = null;
                }
                if (_wifiSelectionMenu != null)
                {
                    _wifiSelectionMenu.Dispose();
                    _wifiSelectionMenu = null;
                }
                if (_suppressedContextMenu != null)
                {
                    _suppressedContextMenu.Dispose();
                    _suppressedContextMenu = null;
                }
                if (_modeNotifyIcon != null)
                {
                    _modeNotifyIcon.Visible = false;
                    _modeNotifyIcon.Dispose();
                    _modeNotifyIcon = null;
                }

                if (_configuredAppIcon != null)
                {
                    _configuredAppIcon.Dispose();
                    _configuredAppIcon = null;
                }
            };
            FormClosing += (_, __) => PersistWindowBounds();
            Resize += (_, __) => PerformResponsiveLayout();
            ResizeEnd += (_, __) => PersistWindowBounds();
            PerformResponsiveLayout();
            UpdateOfficeActionButtonPresentation();
            ApplyTheme();
            SyncExternalThemesAsync();
            if (_settings.WindowMaximized)
            {
                Shown += (_, __) => WindowState = FormWindowState.Maximized;
            }
        }

        private string GetConfiguredAppName()
        {
            var configured = (_settings.SoftwareDisplayName ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(configured) ? AppTitle : configured;
        }

        private string GetConfiguredAppTitle()
        {
            var name = GetConfiguredAppName();
            return name.IndexOf(AppVersion, StringComparison.OrdinalIgnoreCase) >= 0
                ? name
                : string.Format("{0} {1}", name, AppVersion);
        }

        private string GetConfiguredIconPath()
        {
            return ResolveConfiguredPath(_settings.IconFilePath);
        }

        private Icon LoadConfiguredAppIcon()
        {
            var iconPath = GetConfiguredIconPath();
            if (string.IsNullOrWhiteSpace(iconPath) || !File.Exists(iconPath))
            {
                return null;
            }

            try
            {
                var extension = Path.GetExtension(iconPath) ?? string.Empty;
                if (extension.Equals(".ico", StringComparison.OrdinalIgnoreCase))
                {
                    using (var stream = File.OpenRead(iconPath))
                    {
                        return new Icon(stream);
                    }
                }

                if (extension.Equals(".exe", StringComparison.OrdinalIgnoreCase)
                    || extension.Equals(".dll", StringComparison.OrdinalIgnoreCase))
                {
                    var extractedIcon = Icon.ExtractAssociatedIcon(iconPath);
                    return extractedIcon == null ? null : (Icon)extractedIcon.Clone();
                }

                using (var bitmap = new Bitmap(iconPath))
                {
                    var iconHandle = bitmap.GetHicon();
                    try
                    {
                        using (var tempIcon = Icon.FromHandle(iconHandle))
                        {
                            return (Icon)tempIcon.Clone();
                        }
                    }
                    finally
                    {
                        DestroyIcon(iconHandle);
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        private void ApplyConfiguredAppIcon()
        {
            if (_configuredAppIcon != null)
            {
                _configuredAppIcon.Dispose();
                _configuredAppIcon = null;
            }

            _configuredAppIcon = LoadConfiguredAppIcon();
            var iconToUse = _configuredAppIcon ?? SystemIcons.Application;
            Icon = iconToUse;

            if (_modeNotifyIcon != null)
            {
                _modeNotifyIcon.Icon = iconToUse;
            }

            foreach (var window in new[] { _zenModeWindow, _speedrunModeWindow })
            {
                if (window != null && !window.IsDisposed)
                {
                    window.Icon = iconToUse;
                    window.ShowIcon = true;
                }
            }
        }

        private void RefreshAppTitlePresentation()
        {
            var displayTitle = GetConfiguredAppTitle();
            Text = displayTitle;
            if (_titleLabel != null && !_titleLabel.IsDisposed)
            {
                _titleLabel.Text = displayTitle;
            }

            if (_modeNotifyIcon != null)
            {
                _modeNotifyIcon.Text = displayTitle;
            }

            foreach (var pair in new[] { Tuple.Create(_zenModeWindow, "禅模式"), Tuple.Create(_speedrunModeWindow, "速通版") })
            {
                var window = pair.Item1;
                if (window != null && !window.IsDisposed)
                {
                    window.Text = string.Format("{0} - {1}", displayTitle, pair.Item2);
                }
            }

            ApplyConfiguredAppIcon();
        }

        private void ToggleTheme()
        {
            _settings.ThemeMode = _themeButton != null && _themeButton.Checked ? "dark" : "light";
            _settings.Save();
            ApplyTheme();
            SyncExternalThemesAsync();
        }

        private void SettingsButtonMouseUp(object sender, MouseEventArgs e)
        {
            if (e == null || e.Button != MouseButtons.Right)
            {
                return;
            }

            OpenPersonalizationSettings();
        }

        private void ThemeButtonMouseUp(object sender, MouseEventArgs e)
        {
            if (e == null || e.Button != MouseButtons.Right)
            {
                return;
            }

            OpenPersonalizationSettings();
        }

        private void OpenPersonalizationSettings()
        {
            var opened = TryStartTarget("ms-settings:personalization");
            if (!opened)
            {
                opened = TryStartTarget("control.exe", "/name Microsoft.Personalization");
            }

            if (!opened)
            {
                StartTarget("control.exe");
            }
        }

        private void TryOpenStartupMode()
        {
            if (_startupModeOpened)
            {
                return;
            }

            _startupModeOpened = true;
            if (_startupMode == StartupMode.Zen)
            {
                ShowZenModeWindow();
            }
            else if (_startupMode == StartupMode.Speedrun)
            {
                ShowSpeedrunModeWindow();
            }
        }

        private void SyncExternalThemesAsync()
        {
            System.Threading.ThreadPool.QueueUserWorkItem(_ => SyncExternalThemes());
        }

        private void ApplyRememberedWindowBounds()
        {
            try
            {
                if (_settings.WindowWidth <= 0 || _settings.WindowHeight <= 0)
                {
                    return;
                }

                var width = Math.Max(MinimumSize.Width, _settings.WindowWidth);
                var height = Math.Max(MinimumSize.Height, _settings.WindowHeight);
                var targetBounds = new Rectangle(_settings.WindowLeft, _settings.WindowTop, width, height);
                var workingArea = GetRememberedWorkingArea(targetBounds);

                if (workingArea != Rectangle.Empty)
                {
                    width = Math.Min(width, workingArea.Width);
                    height = Math.Min(height, workingArea.Height);
                }

                Size = new Size(width, height);

                if (_settings.WindowLeft != int.MinValue && _settings.WindowTop != int.MinValue)
                {
                    StartPosition = FormStartPosition.Manual;
                    var left = _settings.WindowLeft;
                    var top = _settings.WindowTop;
                    if (workingArea != Rectangle.Empty)
                    {
                        left = Math.Max(workingArea.Left, Math.Min(left, workingArea.Right - width));
                        top = Math.Max(workingArea.Top, Math.Min(top, workingArea.Bottom - height));
                    }

                    Location = new Point(left, top);
                }
            }
            catch
            {
            }
        }

        private void PersistWindowBounds()
        {
            try
            {
                var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
                if (bounds.Width > 0 && bounds.Height > 0)
                {
                    _settings.WindowWidth = Math.Max(300, bounds.Width);
                    _settings.WindowHeight = Math.Max(400, bounds.Height);
                    _settings.WindowLeft = bounds.Left;
                    _settings.WindowTop = bounds.Top;
                }

                _settings.WindowMaximized = WindowState == FormWindowState.Maximized;
                _settings.Save();
            }
            catch
            {
            }
        }

        private static Rectangle GetRememberedWorkingArea(Rectangle targetBounds)
        {
            try
            {
                if (Screen.AllScreens == null || Screen.AllScreens.Length == 0)
                {
                    return Rectangle.Empty;
                }

                if (targetBounds.Width <= 0 || targetBounds.Height <= 0)
                {
                    return Screen.PrimaryScreen == null ? Rectangle.Empty : Screen.PrimaryScreen.WorkingArea;
                }

                foreach (var screen in Screen.AllScreens)
                {
                    if (screen.WorkingArea.IntersectsWith(targetBounds))
                    {
                        return screen.WorkingArea;
                    }
                }

                var center = new Point(
                    targetBounds.Left + Math.Max(1, targetBounds.Width) / 2,
                    targetBounds.Top + Math.Max(1, targetBounds.Height) / 2);
                return Screen.FromPoint(center).WorkingArea;
            }
            catch
            {
                return Screen.PrimaryScreen == null ? Rectangle.Empty : Screen.PrimaryScreen.WorkingArea;
            }
        }

        private Button CreateHeaderButton(string text, EventHandler onClick)
        {
            var button = new Button
            {
                Text = text,
                Size = new Size(68, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(15, 23, 42),
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point),
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            button.Click += onClick;
            return button;
        }

        private void InitializeModeContextMenu()
        {
            _modeContextMenu = new ContextMenuStrip
            {
                ShowImageMargin = false
            };
            _modeContextMenu.Opening += ModeContextMenuOpening;
            _modeContextMenu.Items.Add("显示主窗口", null, (_, __) => RestoreMainFromTray());
            _modeContextMenu.Items.Add(new ToolStripSeparator());
            _modeContextMenu.Items.Add("禅模式", null, (_, __) => ShowZenModeWindow());
            _modeContextMenu.Items.Add("速通版", null, (_, __) => ShowSpeedrunModeWindow());
            _modeContextMenu.Items.Add(new ToolStripSeparator());
            _modeContextMenu.Items.Add("退出", null, (_, __) => ExitEntireApplication());
        }

        private void InitializeOfficeOperationMenu()
        {
            _officeOperationMenu = new ContextMenuStrip
            {
                ShowImageMargin = false
            };
            _officeOperationMenu.Items.Add("复制密钥文件首行", null, (_, __) => SetOfficeKeyOperation("copy"));
            _officeOperationMenu.Items.Add("剪切密钥文件首行", null, (_, __) => SetOfficeKeyOperation("cut"));
            _officeOperationMenu.Items.Add("删除密钥文件首行", null, (_, __) => SetOfficeKeyOperation("delete"));
            _officeOperationMenu.Items.Add(new ToolStripSeparator());
            _officeOperationMenu.Items.Add("打开密钥(源)文件", null, (_, __) => OpenOfficeKeySourceFile());
            _officeOperationMenu.Items.Add("打开桌面(目标)文件", null, (_, __) => OpenOfficeKeyTargetFile());
            _officeSequenceMenuSeparator = new ToolStripSeparator();
            _officeSequenceMenuItem = new ToolStripMenuItem(DefaultOfficeSequenceMenuText, null, (_, __) => ExecuteManualOfficeSequence());
            _officeOperationMenu.Items.Add(_officeSequenceMenuSeparator);
            _officeOperationMenu.Items.Add(_officeSequenceMenuItem);
            _officeOperationMenu.Items.Add(new ToolStripSeparator());
            _officeOperationMenu.Items.Add("重启 Word", null, (_, __) => RestartSpecificOfficeApplication("WINWORD.EXE", "Word"));
            _officeOperationMenu.Items.Add("重启 Excel", null, (_, __) => RestartSpecificOfficeApplication("EXCEL.EXE", "Excel"));
            _officeOperationMenu.Items.Add("重启 PPT", null, (_, __) => RestartSpecificOfficeApplication("POWERPNT.EXE", "PPT"));
            RefreshOfficeSequenceMenuItem();
        }

        private void InitializeWindowsActivationMenu()
        {
            _windowsActivationMenu = new ContextMenuStrip
            {
                ShowImageMargin = false
            };
            _windowsActivationMenu.Items.Add("在 CMD(管理员) 中激活", null, (_, __) => PromptAndRunWindowsActivationCommand());
        }

        private void SetOfficeKeyOperation(string operation)
        {
            _settings.OfficeKeyOperation = NormalizeOfficeKeyOperation(operation);
            _settings.Save();
            UpdateOfficeActionButtonPresentation();
            UpdateOfficeOperationDropButtonPresentation();
        }

        private void ShowWindowsActivationMenu(Control anchorControl)
        {
            if (_windowsActivationMenu == null || anchorControl == null || anchorControl.IsDisposed)
            {
                return;
            }

            var currentSource = _windowsActivationMenu.SourceControl;
            if (_windowsActivationMenu.Visible)
            {
                _windowsActivationMenu.Close(ToolStripDropDownCloseReason.AppClicked);
                if (currentSource == anchorControl)
                {
                    return;
                }
            }

            var preferredSize = _windowsActivationMenu.GetPreferredSize(Size.Empty);
            var screenBounds = Screen.FromControl(anchorControl).WorkingArea;
            var anchorBounds = anchorControl.RectangleToScreen(anchorControl.ClientRectangle);
            Point offset;

            if (anchorBounds.Right + preferredSize.Width + 4 <= screenBounds.Right)
            {
                offset = new Point(anchorControl.Width + 2, 0);
            }
            else if (anchorBounds.Left - preferredSize.Width - 4 >= screenBounds.Left)
            {
                offset = new Point(-preferredSize.Width - 2, 0);
            }
            else
            {
                offset = new Point(0, anchorControl.Height);
            }

            _windowsActivationMenu.Show(anchorControl, offset);
        }

        private void PromptAndRunWindowsActivationCommand()
        {
            using (var dialog = new Form())
            {
                dialog.Text = "Windows 激活";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.ClientSize = new Size(468, 214);
                dialog.Font = Font;

                var label = new Label
                {
                    Text = "请输入 Windows 产品密钥。\r\n支持直接粘贴，程序会自动忽略空格和分隔符。",
                    Location = new Point(18, 18),
                    Size = new Size(430, 40)
                };
                var keyBox = new TextBox
                {
                    Location = new Point(18, 72),
                    Size = new Size(430, 28)
                };
                var tipLabel = new TextBox
                {
                    Text = "确定后会复制设置里的命令模板并打开管理员 CMD；可在设置中修改，支持 {key} 占位符。",
                    Location = new Point(18, 110),
                    Size = new Size(430, 48),
                    ForeColor = Color.DimGray,
                    BackColor = dialog.BackColor,
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    Multiline = true,
                    TabStop = false
                };
                var okButton = new Button { Text = "打开管理员CMD", Location = new Point(242, 164), Size = new Size(110, 32) };
                var cancelButton = new Button { Text = "取消", Location = new Point(364, 164), Size = new Size(84, 32) };

                okButton.Click += (_, __) =>
                {
                    try
                    {
                        var rawInput = (keyBox.Text ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(rawInput))
                        {
                            StartWindowsActivationAdminCommand(null);
                            dialog.DialogResult = DialogResult.OK;
                            dialog.Close();
                            return;
                        }

                        StartWindowsActivationAdminCommand(rawInput);
                        dialog.DialogResult = DialogResult.OK;
                        dialog.Close();
                    }
                    catch (Exception ex)
                    {
                        ShowErrorMessage("打开管理员 CMD 失败：" + ex.Message, "Windows 激活");
                    }
                };
                cancelButton.Click += (_, __) =>
                {
                    dialog.DialogResult = DialogResult.Cancel;
                    dialog.Close();
                };

                dialog.AcceptButton = okButton;
                dialog.CancelButton = cancelButton;
                dialog.Controls.AddRange(new Control[] { label, keyBox, tipLabel, okButton, cancelButton });
                dialog.Shown += (_, __) => keyBox.Focus();
                dialog.ShowDialog(this);
            }
        }

        private void StartWindowsActivationAdminCommand(string normalizedKey)
        {
            var rawKey = (normalizedKey ?? string.Empty).Trim();
            var safeKey = rawKey.Replace("\"", string.Empty);
            var installCommand = BuildWindowsActivationClipboardText(safeKey);
            var displayCommand = EscapeForCmdEcho(installCommand);
            var command = string.Format(
                "title Windows激活 && chcp 65001>nul && echo 请复制下面命令手动执行：&& echo. && echo {0} && echo. && echo 如需联网认证，再执行：slmgr /ato",
                displayCommand);

            try
            {
                try
                {
                    CopyTextToClipboard(installCommand);
                }
                catch
                {
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/k " + command,
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Normal
                };

                Process.Start(startInfo);
                _lastWindowsTimeSyncStatus = string.IsNullOrWhiteSpace(safeKey) ? "已打开管理员 CMD 模板" : "命令已复制并打开管理员 CMD";
                UpdateWindowsActivationButtonPresentation();
            }
            catch (Exception ex)
            {
                ShowErrorMessage("管理员激活执行失败：" + ex.Message, "Windows 激活");
            }
        }

        private static string EscapeForCmdEcho(string value)
        {
            return (value ?? string.Empty)
                .Replace("^", "^^")
                .Replace("&", "^&")
                .Replace("|", "^|")
                .Replace("<", "^<")
                .Replace(">", "^>")
                .Replace("(", "^(")
                .Replace(")", "^)")
                .Replace("%", "%%");
        }

        private void ShowWifiSelectionMenu(Control anchorControl)
        {
            if (anchorControl == null || anchorControl.IsDisposed)
            {
                return;
            }

            if (_wifiSelectionMenu != null)
            {
                _wifiSelectionMenu.Dispose();
                _wifiSelectionMenu = null;
            }

            _wifiSelectionMenu = new ContextMenuStrip
            {
                ShowImageMargin = false
            };

            var orderedProfiles = _settings.WifiProfiles
                .Where(profile => profile != null && !string.IsNullOrWhiteSpace(profile.Ssid))
                .ToList();

            _wifiSelectionMenu.Items.Add("按顺序自动连接", null, (_, __) => ConnectWifiFromSettings(_wifiActionButton));
            _wifiSelectionMenu.Items.Add(new ToolStripSeparator());

            if (orderedProfiles.Count == 0)
            {
                var emptyItem = new ToolStripMenuItem("未设置 WiFi")
                {
                    Enabled = false
                };
                _wifiSelectionMenu.Items.Add(emptyItem);
            }
            else
            {
                for (var i = 0; i < orderedProfiles.Count; i++)
                {
                    var profile = orderedProfiles[i];
                    var profileIndex = i;
                    var label = string.Format(
                        "{0}. {1}",
                        profileIndex + 1,
                        BuildWifiProfileMenuText(profile));
                    _wifiSelectionMenu.Items.Add(label, null, (_, __) => ConnectSpecificWifiProfile(orderedProfiles[profileIndex], _wifiActionButton));
                }
            }

            var currentSource = _wifiSelectionMenu.SourceControl;
            if (_wifiSelectionMenu.Visible)
            {
                _wifiSelectionMenu.Close(ToolStripDropDownCloseReason.AppClicked);
                if (currentSource == anchorControl)
                {
                    return;
                }
            }

            _wifiSelectionMenu.Show(anchorControl, new Point(Math.Max(0, anchorControl.Width / 4), anchorControl.Height));
        }

        private string BuildWifiProfileMenuText(WifiProfile profile)
        {
            if (profile == null)
            {
                return "未知 WiFi";
            }

            var auth = string.Equals(profile.Authentication, "open", StringComparison.OrdinalIgnoreCase)
                ? "开放"
                : MapAuthenticationToDisplay(profile.Authentication);
            return string.Format("{0} [{1}]", profile.Ssid, auth);
        }

        private void ShowOfficeOperationMenu()
        {
            ShowOfficeOperationMenu(_officeOperationDropButton);
        }

        private void ShowOfficeOperationMenu(Control anchorControl)
        {
            if (_officeOperationMenu == null || anchorControl == null || anchorControl.IsDisposed)
            {
                return;
            }

            var currentSource = _officeOperationMenu.SourceControl;
            if (_officeOperationMenu.Visible)
            {
                _officeOperationMenu.Close(ToolStripDropDownCloseReason.AppClicked);
                if (currentSource == anchorControl)
                {
                    return;
                }
            }

            foreach (ToolStripItem item in _officeOperationMenu.Items)
            {
                var menuItem = item as ToolStripMenuItem;
                if (menuItem == null)
                {
                    continue;
                }

                var isOperationChoice = menuItem.Text == "复制密钥文件首行"
                    || menuItem.Text == "剪切密钥文件首行"
                    || menuItem.Text == "删除密钥文件首行";
                if (!isOperationChoice)
                {
                    menuItem.Checked = false;
                    continue;
                }

                menuItem.Checked = string.Equals(
                    MapDisplayToOfficeAction(menuItem.Text),
                    NormalizeOfficeKeyOperation(_settings.OfficeKeyOperation),
                    StringComparison.OrdinalIgnoreCase);
            }

            RefreshOfficeSequenceMenuItem();

            var preferredSize = _officeOperationMenu.GetPreferredSize(Size.Empty);
            var screenBounds = Screen.FromControl(anchorControl).WorkingArea;
            var anchorBounds = anchorControl.RectangleToScreen(anchorControl.ClientRectangle);
            Point offset;

            if (anchorBounds.Right + preferredSize.Width + 4 <= screenBounds.Right)
            {
                offset = new Point(anchorControl.Width + 2, 0);
            }
            else if (anchorBounds.Left - preferredSize.Width - 4 >= screenBounds.Left)
            {
                offset = new Point(-preferredSize.Width - 2, 0);
            }
            else
            {
                offset = new Point(0, anchorControl.Height);
            }

            _officeOperationMenu.Show(anchorControl, offset);
        }

        private void UpdateOfficeOperationDropButtonPresentation()
        {
            if (_officeOperationDropButton == null || _officeOperationDropButton.IsDisposed)
            {
                return;
            }

            _officeOperationDropButton.Visible = false;
            _officeOperationDropButton.Enabled = false;
        }

        private static string NormalizeOfficeKeyOperation(string operation)
        {
            if (string.Equals(operation, "cut", StringComparison.OrdinalIgnoreCase))
            {
                return "cut";
            }

            if (string.Equals(operation, "delete", StringComparison.OrdinalIgnoreCase))
            {
                return "delete";
            }

            return "copy";
        }

        private ContextMenuStrip GetSuppressedContextMenu()
        {
            if (_suppressedContextMenu != null)
            {
                return _suppressedContextMenu;
            }

            _suppressedContextMenu = new ContextMenuStrip
            {
                ShowImageMargin = false
            };
            _suppressedContextMenu.Opening += (_, e) =>
            {
                if (e != null)
                {
                    e.Cancel = true;
                }
            };
            return _suppressedContextMenu;
        }

        private void ModeContextMenuOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (e == null)
            {
                return;
            }

            if (_modeNotifyIcon != null && _modeNotifyIcon.ContextMenuStrip == sender)
            {
                return;
            }

            var sourceControl = _modeContextMenu == null ? null : _modeContextMenu.SourceControl;
            if (ShouldSuppressModeContextMenu(sourceControl))
            {
                e.Cancel = true;
            }
        }

        private bool ShouldSuppressModeContextMenu(Control sourceControl)
        {
            if (sourceControl == null || sourceControl.IsDisposed)
            {
                return false;
            }

            var localPoint = sourceControl.PointToClient(Control.MousePosition);
            var deepest = GetDeepestChildAtPoint(sourceControl, localPoint);
            if (deepest is Button || deepest is CheckBox || deepest is TextBox)
            {
                return true;
            }

            return false;
        }

        private static Control GetDeepestChildAtPoint(Control root, Point clientPoint)
        {
            var current = root;
            var point = clientPoint;
            while (current != null)
            {
                var child = current.GetChildAtPoint(point, GetChildAtPointSkip.Invisible);
                if (child == null)
                {
                    return current;
                }

                point = child.PointToClient(current.PointToScreen(point));
                current = child;
            }

            return root;
        }

        private void AttachModeContextMenu(Control root)
        {
            if (root == null || _modeContextMenu == null)
            {
                return;
            }

            if (root is Button || root is CheckBox || root is TextBox)
            {
                root.ContextMenuStrip = GetSuppressedContextMenu();
            }
            else
            {
                root.ContextMenuStrip = _modeContextMenu;
            }

            foreach (Control child in root.Controls)
            {
                AttachModeContextMenu(child);
            }
        }

        private void ShowZenModeWindow()
        {
            _zenModeWindow = ShowCompactModeWindow(_zenModeWindow, "禅模式", (returnToMain, exitApplication, openSettings) => CreateZenModeWindowButtons(returnToMain, exitApplication, openSettings), () => _zenModeWindow = null);
        }

        private void ShowSpeedrunModeWindow()
        {
            _speedrunModeWindow = ShowCompactModeWindow(_speedrunModeWindow, "速通版", (returnToMain, exitApplication, openSettings) => CreateSpeedrunModeWindowButtons(returnToMain, exitApplication, openSettings), () => _speedrunModeWindow = null);
        }

        private Form ShowCompactModeWindow(Form existingWindow, string title, Func<Action, Action, Action, Button[]> buttonFactory, Action onClosed)
        {
            if (existingWindow != null && !existingWindow.IsDisposed)
            {
                existingWindow.TopMost = TopMost;
                existingWindow.Activate();
                HideMainToTray(title);
                return existingWindow;
            }

            var restoreMainOnClose = false;
            var exitApplicationOnClose = false;
            var openSettingsOnClose = false;
            Form window = null;
            Action returnToMain = () =>
            {
                restoreMainOnClose = true;
                CloseCompactWindow(window);
            };
            Action exitApplication = () =>
            {
                exitApplicationOnClose = true;
                CloseCompactWindow(window);
            };
            Action openSettings = () =>
            {
                restoreMainOnClose = true;
                openSettingsOnClose = true;
                CloseCompactWindow(window);
            };

            window = new Form
            {
                Text = string.Format("{0} - {1}", GetConfiguredAppTitle(), title),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedSingle,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowIcon = true,
                ShowInTaskbar = false,
                AutoScroll = false,
                ClientSize = new Size(360, 420),
                MinimumSize = new Size(360, 420),
                MaximumSize = new Size(360, 420),
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                TopMost = TopMost
            };
            window.Icon = _configuredAppIcon ?? SystemIcons.Application;

            window.Location = GetCompactModeWindowLocation(title, window.Size);
            var buttons = buttonFactory == null ? new Button[0] : buttonFactory(returnToMain, exitApplication, openSettings);
            LayoutCompactModeButtons(window, buttons ?? new Button[0]);
            AddCompactOfficeOperationDropButton(window);
            ApplyModeWindowTheme(window);
            window.Move += (_, __) => PersistCompactModeWindowLocation(title, window);
            window.FormClosed += (_, __) =>
            {
                PersistCompactModeWindowLocation(title, window);
                if (onClosed != null)
                {
                    onClosed();
                }

                if (exitApplicationOnClose || _applicationExiting)
                {
                    ExitEntireApplication();
                }
                else if (restoreMainOnClose && !HasCompactModeWindowOpen())
                {
                    RestoreMainFromTray();
                    if (openSettingsOnClose && !IsDisposed)
                    {
                        BeginInvoke((MethodInvoker)delegate
                        {
                            ShowSettingsDialog();
                        });
                    }
                }
            };
            window.Show();
            HideMainToTray(title);
            return window;
        }

        private Point GetCompactModeWindowLocation(string title, Size windowSize)
        {
            var workingArea = Screen.PrimaryScreen == null ? Rectangle.Empty : Screen.PrimaryScreen.WorkingArea;
            if (workingArea == Rectangle.Empty)
            {
                return new Point(100, 100);
            }

            int savedLeft;
            int savedTop;
            if (TryGetCompactModeWindowLocation(title, out savedLeft, out savedTop))
            {
                var savedBounds = new Rectangle(savedLeft, savedTop, Math.Max(360, windowSize.Width), Math.Max(420, windowSize.Height));
                var savedWorkingArea = GetRememberedWorkingArea(savedBounds);
                if (savedWorkingArea != Rectangle.Empty)
                {
                    var xSaved = Math.Max(savedWorkingArea.Left, Math.Min(savedLeft, savedWorkingArea.Right - Math.Max(360, windowSize.Width)));
                    var ySaved = Math.Max(savedWorkingArea.Top, Math.Min(savedTop, savedWorkingArea.Bottom - Math.Max(420, windowSize.Height)));
                    return new Point(xSaved, ySaved);
                }
            }

            var x = Math.Max(workingArea.Left, workingArea.Right - Math.Max(360, windowSize.Width) - 24);
            var y = Math.Max(workingArea.Top + 36, workingArea.Top + (workingArea.Height - Math.Max(420, windowSize.Height)) / 2);
            return new Point(x, y);
        }

        private bool TryGetCompactModeWindowLocation(string title, out int left, out int top)
        {
            left = int.MinValue;
            top = int.MinValue;
            var isZen = title != null && title.IndexOf("禅", StringComparison.OrdinalIgnoreCase) >= 0;
            if (isZen)
            {
                left = _settings.ZenWindowLeft;
                top = _settings.ZenWindowTop;
            }
            else
            {
                left = _settings.SpeedrunWindowLeft;
                top = _settings.SpeedrunWindowTop;
            }

            return left != int.MinValue && top != int.MinValue;
        }

        private void PersistCompactModeWindowLocation(string title, Form window)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            try
            {
                var isZen = title != null && title.IndexOf("禅", StringComparison.OrdinalIgnoreCase) >= 0;
                if (isZen)
                {
                    _settings.ZenWindowLeft = window.Left;
                    _settings.ZenWindowTop = window.Top;
                }
                else
                {
                    _settings.SpeedrunWindowLeft = window.Left;
                    _settings.SpeedrunWindowTop = window.Top;
                }

                _settings.Save();
            }
            catch
            {
            }
        }

        private void ExitEntireApplication()
        {
            if (_applicationExiting)
            {
                if (!IsDisposed)
                {
                    Close();
                }
                return;
            }

            _applicationExiting = true;
            if (_modeNotifyIcon != null)
            {
                _modeNotifyIcon.Visible = false;
            }

            foreach (var window in new[] { _zenModeWindow, _speedrunModeWindow })
            {
                if (window != null && !window.IsDisposed)
                {
                    window.Close();
                }
            }

            if (!IsDisposed)
            {
                Close();
            }
        }

        private Button[] CreateZenModeWindowButtons(Action returnToMain, Action exitApplication, Action openSettings)
        {
            return new[]
            {
                CreateCompactModeButton(GetCompactWifiButtonText(), (_, __) => ConnectWifiFromSettings(_wifiActionButton), BuildWifiActionTooltipText(), "__compactWifiButton", _wifiActionButton),
                CreateCompactModeButton("Windows\r\n激活", (_, __) => OpenWindowsActivation(), "左击：打开系统激活页并尝试校准系统时间；右击：显示管理员 CMD 激活菜单。", "__compactWindowsButton", _windowsActionButton),
                CreateCompactModeButton("配置", (_, __) => OpenConfigurationCheck(), "先打开设置的系统-关于，再打开磁盘管理。", null, _configActionButton),
                CreateCompactModeButton(GetCompactBatteryButtonText(), (_, __) => OpenBatteryCheck(), BuildCompactBatteryButtonTooltipText(), "__compactBatteryButton", _batteryActionButton),
                CreateCompactModeButton("摄像头", (_, __) => OpenCameraCheck(), "打开相机检查画面", null, _cameraActionButton),
                CreateCompactModeButton("键盘", (_, __) => OpenKeyboardCheck(), "启动键盘测试工具", null, _keyboardActionButton),
                CreateCompactModeButton(GetCompactOfficeButtonText(), (_, __) => OpenOfficeActivation(_officeActionButton), BuildOfficeActionTooltipText(), "__compactOfficeButton", _officeActionButton),
                CreateCompactModeButton("擦除", (_, __) => ClearHistoryAndRecycleBin(_clearActionButton), "清理最近记录、浏览器历史与回收站", null, _clearActionButton),
                CreateCompactModeButton("关机/重启", (_, __) => ShowPowerDialog(), "显示重启或关机确认窗口"),
                CreateCompactModeButton("设置", (_, __) => openSettings(), "关闭禅模式窗口并打开设置", "__compactSettingsButton"),
                CreateCompactModeButton("支持作者", (_, __) => OpenConfiguredWebLink(_settings.DonateUrl, "支持作者"), "打开支持作者链接", "__compactSupportButton"),
                CreateCompactModeButton("扬声器测试", (_, __) => OpenSpeakerTest(), "播放扬声器测试资源并清理该文件最近记录", "__compactSpeakerButton"),
                CreateCompactModeButton(GetCompactThemeButtonText(), (_, __) => ToggleThemeFromCompactWindow(), "切换明亮/黑暗主题", "__compactThemeButton"),
                CreateCompactModeButton("返回主窗口", (_, __) => returnToMain(), "关闭禅模式窗口并回到主窗口", "__compactReturnButton"),
                CreateCompactModeButton("退出程序", (_, __) => exitApplication(), "退出整个应用，包括主窗口", "__compactExitButton")
            };
        }

        private Button[] CreateSpeedrunModeWindowButtons(Action returnToMain, Action exitApplication, Action openSettings)
        {
            return new[]
            {
                CreateCompactModeButton("Windows\r\n激活", (_, __) => OpenWindowsActivation(), "左击：打开系统激活页并尝试校准系统时间；右击：显示管理员 CMD 激活菜单。", "__compactWindowsButton", _windowsActionButton),
                CreateCompactModeButton("配置", (_, __) => OpenConfigurationCheck(), "先打开设置的系统-关于，再打开磁盘管理。", null, _configActionButton),
                CreateCompactModeButton(GetCompactBatteryButtonText(), (_, __) => OpenBatteryCheck(), BuildCompactBatteryButtonTooltipText(), "__compactBatteryButton", _batteryActionButton),
                CreateCompactModeButton(GetCompactOfficeButtonText(), (_, __) => OpenOfficeActivation(_officeActionButton), BuildOfficeActionTooltipText(), "__compactOfficeButton", _officeActionButton),
                CreateCompactModeButton("设置", (_, __) => openSettings(), "关闭速通版窗口并打开设置", "__compactSettingsButton"),
                CreateCompactModeButton("支持作者", (_, __) => OpenConfiguredWebLink(_settings.DonateUrl, "支持作者"), "打开支持作者链接", "__compactSupportButton"),
                CreateCompactModeButton(GetCompactThemeButtonText(), (_, __) => ToggleThemeFromCompactWindow(), "切换明亮/黑暗主题", "__compactThemeButton"),
                CreateCompactModeButton("返回主窗口", (_, __) => returnToMain(), "关闭速通版窗口并回到主窗口", "__compactReturnButton"),
                CreateCompactModeButton("退出程序", (_, __) => exitApplication(), "退出整个应用，包括主窗口", "__compactExitButton")
            };
        }

        private Button CreateCompactModeButton(string text, EventHandler onClick, string toolTipText, string name = null, Button sourceButton = null)
        {
            var button = new Button
            {
                Name = string.IsNullOrWhiteSpace(name) ? string.Empty : name,
                Text = text,
                Size = new Size(120, 48),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point),
                TextAlign = ContentAlignment.MiddleCenter
            };
            if (IsCompactFooterLinkButtonName(name))
            {
                button.Font = new Font("Microsoft YaHei UI", 7.0F, FontStyle.Underline, GraphicsUnit.Point);
                button.FlatAppearance.BorderSize = 0;
                button.BackColor = BackColor;
                button.Size = new Size(54, 22);
                button.TabStop = false;
                button.Cursor = Cursors.Hand;
                button.Padding = new Padding(0, 0, 0, 1);
            }
            else
            {
                button.FlatAppearance.BorderSize = 1;
            }
            var sourceInfo = sourceButton == null ? null : sourceButton.Tag as ButtonStateInfo;
            if (sourceInfo != null)
            {
                button.Tag = new ButtonStateInfo
                {
                    BaseTitle = text,
                    Subtitle = string.Empty,
                    State = sourceInfo.State,
                    TrackCompletion = sourceInfo.TrackCompletion,
                    LinkedButton = sourceButton,
                    CompactModeStyle = true
                };
                ApplyButtonVisualState(button);
                button.MouseDown += ActionButtonMouseDown;
            }

            button.Click += (_, e) =>
            {
                onClick(button, e);
                if (sourceButton != null && !sourceButton.IsDisposed)
                {
                    var linkedInfo = sourceButton.Tag as ButtonStateInfo;
                    if (linkedInfo != null && linkedInfo.TrackCompletion && linkedInfo.State == "default")
                    {
                        SetButtonState(sourceButton, "clicked");
                    }
                    SyncCompactModeButtonStates();
                }
            };
            if (!string.IsNullOrWhiteSpace(toolTipText))
            {
                _toolTip.SetToolTip(button, toolTipText);
            }

            if (string.Equals(name, "__compactThemeButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactSettingsButton", StringComparison.Ordinal))
            {
                button.MouseUp += CompactPersonalizationButtonMouseUp;
                _toolTip.SetToolTip(button, "左击执行当前功能；右击打开系统个性化");
            }

            return button;
        }

        private void AddCompactOfficeOperationDropButton(Form window)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            foreach (var menuButton in window.Controls.OfType<Button>()
                .Where(button => string.Equals(button.Name, "__compactOfficeMenuButton", StringComparison.Ordinal))
                .ToArray())
            {
                window.Controls.Remove(menuButton);
                menuButton.Dispose();
            }
        }

        private void LayoutCompactModeButtons(Form window, IList<Button> buttons)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            window.SuspendLayout();
            window.Controls.Clear();

            var gridButtons = buttons.Where(button => button != null && !IsCompactFooterLinkButton(button)).ToList();
            var footerLinkButtons = buttons.Where(IsCompactFooterLinkButton).ToList();
            var columns = 3;
            const int horizontalGap = 10;
            var verticalGap = gridButtons.Count > 10 ? 6 : 8;
            const int topMargin = 14;
            var buttonWidth = Math.Max(84, (window.ClientSize.Width - 24 - horizontalGap * (columns - 1)) / columns);
            var buttonHeight = gridButtons.Count > 10 ? 44 : 48;
            var totalWidth = columns * buttonWidth + (columns - 1) * horizontalGap;
            var startX = Math.Max(12, (window.ClientSize.Width - totalWidth) / 2);
            var contentBottom = topMargin;

            for (var index = 0; index < gridButtons.Count; index++)
            {
                var row = index / columns;
                var column = index % columns;
                var button = gridButtons[index];
                button.Size = new Size(buttonWidth, buttonHeight);
                button.Location = new Point(
                    startX + column * (buttonWidth + horizontalGap),
                    topMargin + row * (buttonHeight + verticalGap));
                window.Controls.Add(button);
                contentBottom = Math.Max(contentBottom, button.Bottom);
            }

            if (footerLinkButtons.Count > 0)
            {
                const int linkTopGap = 6;
                const int linkGap = 6;
                const int linkRowGap = 4;
                const int linkHeight = 22;
                var linkY = contentBottom + linkTopGap;
                var footerColumns = Math.Min(3, footerLinkButtons.Count);
                var footerWidth = Math.Max(72, (window.ClientSize.Width - 24 - (footerColumns - 1) * linkGap) / footerColumns);
                for (var index = 0; index < footerLinkButtons.Count; index++)
                {
                    var row = index / footerColumns;
                    var column = index % footerColumns;
                    var linkX = 12 + column * (footerWidth + linkGap);
                    var button = footerLinkButtons[index];
                    button.Size = new Size(footerWidth, linkHeight);
                    button.Location = new Point(linkX, linkY + row * (linkHeight + linkRowGap));
                    window.Controls.Add(button);
                    contentBottom = Math.Max(contentBottom, button.Bottom);
                }
            }

            AddOrUpdateCompactModeSummary(window, contentBottom);
            window.ResumeLayout(false);
        }

        private static bool IsCompactFooterLinkButton(Button button)
        {
            return button != null && IsCompactFooterLinkButtonName(button.Name);
        }

        private static bool IsCompactFooterLinkButtonName(string name)
        {
            return string.Equals(name, "__compactSettingsButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactSupportButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactThemeButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactSpeakerButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactReturnButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactExitButton", StringComparison.Ordinal);
        }

        private static bool IsCompactBlueLinkButtonName(string name)
        {
            return string.Equals(name, "__compactThemeButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactSettingsButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactSupportButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactSpeakerButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactReturnButton", StringComparison.Ordinal)
                || string.Equals(name, "__compactExitButton", StringComparison.Ordinal);
        }

        private void AddOrUpdateCompactModeSummary(Form window, int contentBottom)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            const int sideMargin = 12;
            const int topGap = 6;
            const int bottomGap = 8;
            var panelTop = Math.Max(contentBottom + topGap, 226);
            var panelHeight = Math.Max(72, window.ClientSize.Height - panelTop - bottomGap);

            var summaryPanel = window.Controls.OfType<Panel>()
                .FirstOrDefault(panel => string.Equals(panel.Name, "__compactSummaryPanel", StringComparison.Ordinal));
            if (summaryPanel == null)
            {
                summaryPanel = new Panel
                {
                    Name = "__compactSummaryPanel"
                };

                var bodyBox = new RichTextBox
                {
                    Name = "__compactSummaryBody",
                    ReadOnly = true,
                    TabStop = false,
                    ShortcutsEnabled = true,
                    BorderStyle = BorderStyle.None,
                    Font = new Font("Microsoft YaHei UI", 8.2F, FontStyle.Regular, GraphicsUnit.Point),
                    ScrollBars = RichTextBoxScrollBars.Vertical,
                    DetectUrls = false
                };

                summaryPanel.Controls.Add(bodyBox);
                window.Controls.Add(summaryPanel);
            }

            summaryPanel.Location = new Point(sideMargin, panelTop);
            summaryPanel.Size = new Size(window.ClientSize.Width - sideMargin * 2, panelHeight);

            var summaryBody = summaryPanel.Controls["__compactSummaryBody"] as RichTextBox;
            if (summaryBody != null)
            {
                summaryBody.Location = new Point(10, 10);
                summaryBody.Size = new Size(summaryPanel.Width - 20, Math.Max(52, summaryPanel.Height - 20));
                RenderCompactModeSummary(summaryBody);
                summaryBody.SelectionLength = 0;
            }
        }

        private void UpdateCompactModeSummaries()
        {
            UpdateCompactModeSummary(_zenModeWindow);
            UpdateCompactModeSummary(_speedrunModeWindow);
        }

        private void UpdateCompactModeSummary(Form window)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            var summaryPanel = window.Controls.OfType<Panel>()
                .FirstOrDefault(panel => string.Equals(panel.Name, "__compactSummaryPanel", StringComparison.Ordinal));
            if (summaryPanel == null)
            {
                return;
            }

            var summaryBody = summaryPanel.Controls["__compactSummaryBody"] as RichTextBox;
            if (summaryBody == null)
            {
                return;
            }

            RenderCompactModeSummary(summaryBody);
            summaryBody.SelectionLength = 0;
        }

        private static void ApplyCompactModeSummaryTheme(Form window, Color panelBack, Color bodyText)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            var summaryPanel = window.Controls.OfType<Panel>()
                .FirstOrDefault(panel => string.Equals(panel.Name, "__compactSummaryPanel", StringComparison.Ordinal));
            if (summaryPanel == null)
            {
                return;
            }

            summaryPanel.BackColor = panelBack;

            var summaryBody = summaryPanel.Controls["__compactSummaryBody"] as RichTextBox;
            if (summaryBody != null)
            {
                summaryBody.BackColor = panelBack;
            }
        }

        private void RenderCompactModeSummary(RichTextBox summaryBody)
        {
            if (summaryBody == null || summaryBody.IsDisposed)
            {
                return;
            }

            var dark = IsDarkTheme();
            var labelColor = dark ? Color.FromArgb(96, 165, 250) : Color.FromArgb(29, 78, 216);
            var textColor = dark ? Color.FromArgb(226, 232, 240) : Color.FromArgb(15, 23, 42);
            var noteColor = dark ? Color.FromArgb(148, 163, 184) : Color.FromArgb(71, 85, 105);
            var successColor = dark ? Color.FromArgb(134, 239, 172) : Color.FromArgb(21, 128, 61);
            var warnColor = dark ? Color.FromArgb(253, 224, 71) : Color.FromArgb(180, 83, 9);
            var failColor = dark ? Color.FromArgb(252, 165, 165) : Color.FromArgb(185, 28, 28);

            summaryBody.SuspendLayout();
            summaryBody.ForeColor = textColor;
            summaryBody.Clear();

            foreach (var item in BuildCompactModeSummaryItems())
            {
                var valueColor = textColor;
                if (string.Equals(item.Key, "说明", StringComparison.Ordinal))
                {
                    valueColor = noteColor;
                }
                else if (string.Equals(item.Key, "电池", StringComparison.Ordinal))
                {
                    valueColor = item.Value.IndexOf("无电池", StringComparison.OrdinalIgnoreCase) >= 0 ? failColor : successColor;
                }
                else if (string.Equals(item.Key, "蓝牙", StringComparison.Ordinal))
                {
                    valueColor = string.Equals(item.Value, "有", StringComparison.Ordinal) ? successColor : failColor;
                }
                else if (string.Equals(item.Key, "WiFi", StringComparison.Ordinal))
                {
                    valueColor = item.Value.IndexOf("未", StringComparison.OrdinalIgnoreCase) >= 0 ? failColor : successColor;
                }
                else if (string.Equals(item.Key, "密码", StringComparison.Ordinal))
                {
                    valueColor = string.Equals(item.Value, "无密码", StringComparison.Ordinal) ? successColor : warnColor;
                }
                else if (string.Equals(item.Key, "Office", StringComparison.Ordinal))
                {
                    valueColor = GetAvailableOfficeKeyCount() > 0 ? warnColor : failColor;
                }

                AppendCompactModeSummaryLine(summaryBody, item.Key, item.Value, labelColor, valueColor);
            }

            summaryBody.SelectionStart = 0;
            summaryBody.SelectionLength = 0;
            summaryBody.ResumeLayout();
        }

        private static void AppendCompactModeSummaryLine(RichTextBox summaryBody, string label, string value, Color labelColor, Color valueColor)
        {
            if (summaryBody.TextLength > 0)
            {
                summaryBody.AppendText(Environment.NewLine);
            }

            summaryBody.SelectionColor = labelColor;
            summaryBody.AppendText(label + " ");
            summaryBody.SelectionColor = valueColor;
            summaryBody.AppendText(value ?? string.Empty);
            summaryBody.SelectionColor = summaryBody.ForeColor;
        }

        private KeyValuePair<string, string>[] BuildCompactModeSummaryItems()
        {
            return new[]
            {
                new KeyValuePair<string, string>("机型", BuildCompactModelSummaryText()),
                new KeyValuePair<string, string>("系统", BuildCompactOsSummaryText()),
                new KeyValuePair<string, string>("配置", BuildCompactConfigurationSummaryText()),
                new KeyValuePair<string, string>("CPU", BuildCompactCpuSummaryText()),
                new KeyValuePair<string, string>("电池", BuildCompactBatterySummaryText()),
                new KeyValuePair<string, string>("蓝牙", BuildCompactBluetoothSummaryText()),
                new KeyValuePair<string, string>("WiFi", BuildCompactWifiSummaryText()),
                new KeyValuePair<string, string>("密码", BuildCompactWifiPasswordSummaryText()),
                new KeyValuePair<string, string>("Office", string.Format("余{0}码", GetAvailableOfficeKeyCount())),
                new KeyValuePair<string, string>("说明", "返回主窗口=回主窗 / 退出程序=退程序 / X=留托盘")
            };
        }

        private string BuildCompactModeSummaryText()
        {
            return string.Join(Environment.NewLine, BuildCompactModeSummaryItems().Select(item => item.Key + " " + item.Value).ToArray());
        }

        private string BuildCompactModelSummaryText()
        {
            var modelText = NormalizeSingleLineText(_modelLabel == null ? string.Empty : _modelLabel.Text);
            if (string.IsNullOrWhiteSpace(modelText) || string.Equals(modelText, "无数据", StringComparison.Ordinal))
            {
                return "无数据";
            }

            return ShortenDisplayText(modelText, 44);
        }

        private string BuildCompactOsSummaryText()
        {
            var osText = NormalizeSingleLineText(_osLabel == null ? string.Empty : _osLabel.Text);
            if (string.IsNullOrWhiteSpace(osText) || string.Equals(osText, "无数据", StringComparison.Ordinal))
            {
                return "无数据";
            }

            return ShortenDisplayText(osText, 30);
        }

        private string BuildCompactConfigurationSummaryText()
        {
            var memoryText = ExtractCompactCapacityText(_memoryLabel == null ? string.Empty : _memoryLabel.Text, false);
            var diskText = ExtractCompactCapacityText(_diskLabel == null ? string.Empty : _diskLabel.Text, true);

            if (!string.IsNullOrWhiteSpace(memoryText) && !string.IsNullOrWhiteSpace(diskText))
            {
                return string.Format("{0}+{1}", memoryText, diskText);
            }

            if (!string.IsNullOrWhiteSpace(memoryText))
            {
                return memoryText;
            }

            if (!string.IsNullOrWhiteSpace(diskText))
            {
                return diskText;
            }

            return "无数据";
        }

        private string BuildCompactCpuSummaryText()
        {
            var cpuText = NormalizeSingleLineText(_cpuLabel == null ? string.Empty : _cpuLabel.Text);
            if (string.IsNullOrWhiteSpace(cpuText) || string.Equals(cpuText, "无数据", StringComparison.Ordinal))
            {
                return "无数据";
            }

            var upperCpu = cpuText.ToUpperInvariant();
            var intelMatch = Regex.Match(upperCpu, @"(?:CORE\s+)?I([3579])[-\s]?(\d{4,5})");
            if (intelMatch.Success)
            {
                var tier = intelMatch.Groups[1].Value;
                var modelNumber = intelMatch.Groups[2].Value;
                var generation = modelNumber.Length >= 5 ? modelNumber.Substring(0, 2) : modelNumber.Substring(0, 1);
                return string.Format("i{0}-{1}th", tier, generation);
            }

            var ryzenMatch = Regex.Match(upperCpu, @"RYZEN\s+([3579])\s+(\d{4,5}[A-Z]{0,2})");
            if (ryzenMatch.Success)
            {
                return string.Format("R{0}-{1}", ryzenMatch.Groups[1].Value, ryzenMatch.Groups[2].Value);
            }

            var nSeriesMatch = Regex.Match(upperCpu, @"\bN(\d{3,4})\b");
            if (nSeriesMatch.Success)
            {
                return string.Format("N{0}", nSeriesMatch.Groups[1].Value);
            }

            return ShortenDisplayText(cpuText, 22);
        }

        private string BuildCompactBatterySummaryText()
        {
            var statusText = NormalizeSingleLineText(_batteryStatusLabel == null ? string.Empty : _batteryStatusLabel.Text);
            if (!_batteryPresent || _batteryHealthPercent < 0)
            {
                if (statusText.IndexOf("未检测到电池", StringComparison.OrdinalIgnoreCase) >= 0
                    || statusText.IndexOf("无电池", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "无电池 / 剩余无 / 健康无";
                }

                return string.IsNullOrWhiteSpace(statusText) ? "无数据" : ShortenDisplayText(statusText, 22);
            }

            var capacityText = FormatCompactBatteryCapacity(_batteryFullChargeCapacityMWh);
            var remainingText = _batteryRemainingPercent >= 0
                ? string.Format("{0:0}%", _batteryRemainingPercent)
                : (string.IsNullOrWhiteSpace(capacityText) ? "无数据" : capacityText);
            return string.Format("有电池 / 剩余{0} / 健康{1:0.#}%", remainingText, _batteryHealthPercent);
        }

        private string BuildCompactWifiSummaryText()
        {
            bool connected;
            string connectedSsid;
            if (TryGetCurrentWifiConnection(out connected, out connectedSsid) && connected)
            {
                return ShortenDisplayText(connectedSsid, 16);
            }

            return _settings.WifiProfiles.Count == 0 ? "未设置" : "未连接";
        }

        private string BuildCompactWifiPasswordSummaryText()
        {
            WifiProfile matchedProfile = null;
            bool connected;
            string connectedSsid;
            if (TryGetCurrentWifiConnection(out connected, out connectedSsid) && connected)
            {
                matchedProfile = _settings.WifiProfiles.FirstOrDefault(profile =>
                    profile != null && string.Equals(profile.Ssid, connectedSsid, StringComparison.OrdinalIgnoreCase));
            }

            if (matchedProfile == null)
            {
                matchedProfile = _settings.WifiProfiles.FirstOrDefault(profile => profile != null && !string.IsNullOrWhiteSpace(profile.Ssid));
            }

            if (matchedProfile == null)
            {
                return "未设置";
            }

            if (string.Equals(matchedProfile.Authentication, "open", StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(matchedProfile.Password))
            {
                return "无密码";
            }

            return ShortenDisplayText(matchedProfile.Password, 24);
        }

        private string BuildCompactBluetoothSummaryText()
        {
            var bluetoothText = NormalizeSingleLineText(_bluetoothLabel == null ? string.Empty : _bluetoothLabel.Text);
            if (string.IsNullOrWhiteSpace(bluetoothText))
            {
                return "无数据";
            }

            if (bluetoothText.IndexOf("未检测到", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "无";
            }

            if (bluetoothText.IndexOf("检测失败", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "检测失败";
            }

            return "有";
        }

        private static string ExtractCompactCapacityText(string text, bool isDisk)
        {
            var normalizedText = NormalizeSingleLineText(text);
            if (string.IsNullOrWhiteSpace(normalizedText) || string.Equals(normalizedText, "无数据", StringComparison.Ordinal))
            {
                return string.Empty;
            }

            if (isDisk && normalizedText.StartsWith("总容量", StringComparison.Ordinal))
            {
                normalizedText = normalizedText.Substring("总容量".Length).Trim();
            }

            normalizedText = Regex.Replace(normalizedText, @"（.*?）", string.Empty).Trim();
            var match = Regex.Match(normalizedText, @"(\d+(?:\.\d+)?)\s*(TB|GB|MB)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return ShortenDisplayText(normalizedText, 14);
            }

            return FormatCompactCapacity(match.Groups[1].Value, match.Groups[2].Value);
        }

        private static string FormatCompactCapacity(string numberText, string unitText)
        {
            double value;
            if (!double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                && !double.TryParse(numberText, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
            {
                return string.Format("{0}{1}", numberText, unitText == null ? string.Empty : unitText.Trim().ToUpperInvariant());
            }

            var suffix = (unitText ?? string.Empty).Trim().ToUpperInvariant();
            if (Math.Abs(value - Math.Round(value)) < 0.05d)
            {
                return string.Format("{0:0}{1}", Math.Round(value), suffix);
            }

            return string.Format("{0:0.#}{1}", value, suffix);
        }

        private static string FormatCompactBatteryCapacity(double milliWattHours)
        {
            if (milliWattHours <= 0)
            {
                return string.Empty;
            }

            var wattHours = milliWattHours / 1000d;
            return wattHours >= 100d
                ? string.Format("{0:0}Wh", wattHours)
                : string.Format("{0:0.#}Wh", wattHours);
        }

        private static string NormalizeSingleLineText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var firstLine = (text ?? string.Empty)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .FirstOrDefault();
            if (string.IsNullOrWhiteSpace(firstLine))
            {
                return string.Empty;
            }

            return Regex.Replace(firstLine.Trim(), @"\s+", " ");
        }

        private static string ShortenDisplayText(string text, int maxLength)
        {
            var normalized = NormalizeSingleLineText(text);
            if (string.IsNullOrWhiteSpace(normalized) || maxLength <= 0 || normalized.Length <= maxLength)
            {
                return normalized;
            }

            return normalized.Substring(0, Math.Max(1, maxLength - 3)).TrimEnd() + "...";
        }

        private int GetAvailableOfficeKeyCount()
        {
            var sourcePath = BuildPortableCandidates(_settings.OfficeKeySourcePath, OfficeKeySourcePath).FirstOrDefault(File.Exists);
            return CountAvailableOfficeProductKeys(sourcePath);
        }

        private string BuildCompactBatteryButtonTooltipText()
        {
            var summary = BuildCompactBatterySummaryText();
            return string.IsNullOrWhiteSpace(summary)
                ? "左击：打开 BatteryInfoView。右击：手动判通过或恢复真实检测。"
                : summary + "\r\n左击：打开 BatteryInfoView。右击：手动判通过或恢复真实检测。";
        }

        private void ToggleThemeFromCompactWindow()
        {
            if (_themeButton != null && !_themeButton.IsDisposed)
            {
                _themeButton.Checked = !_themeButton.Checked;
                return;
            }

            ToggleTheme();
        }

        private void CompactPersonalizationButtonMouseUp(object sender, MouseEventArgs e)
        {
            if (e == null || e.Button != MouseButtons.Right)
            {
                return;
            }

            OpenPersonalizationSettings();
        }

        private static void CloseCompactWindow(Form window)
        {
            if (window != null && !window.IsDisposed)
            {
                window.Close();
            }
        }

        private string GetCompactOfficeButtonText()
        {
            return string.Format(
                "Office余{0}码\r\n{1}",
                GetAvailableOfficeKeyCount(),
                BuildCurrentOfficeActionCountText());
        }

        private string GetCompactBatteryButtonText()
        {
            if (!_batteryPresent || _batteryHealthPercent < 0)
            {
                return "电池\r\n无电池";
            }

            return string.Format("电池\r\n健康{0:0.#}%", _batteryHealthPercent);
        }

        private string GetCompactWifiButtonText()
        {
            bool connected;
            string connectedSsid;
            if (TryGetCurrentWifiConnection(out connected, out connectedSsid) && connected)
            {
                return string.Format("联网\r\n{0}", ShortenDisplayText(connectedSsid, 10));
            }

            if (_settings.WifiProfiles.Count == 0)
            {
                return "联网\r\n未设置";
            }

            return "联网\r\n未连接";
        }

        private string GetCompactThemeButtonText()
        {
            return IsDarkTheme() ? "切亮主题" : "切暗主题";
        }

        private string BuildOfficeActionSubtitle()
        {
            var normalizedOperation = NormalizeOfficeKeyOperation(_settings.OfficeKeyOperation);
            var countText = string.Format("余{0}码 {1}", GetAvailableOfficeKeyCount(), BuildCurrentOfficeActionCountText());
            if (string.Equals(normalizedOperation, "delete", StringComparison.OrdinalIgnoreCase))
            {
                return string.Format("{0}\r\n仅删除首行", countText);
            }

            return string.Format("{0}\r\n点按处理并打开", countText);
        }

        private string BuildOfficeActionTooltipText()
        {
            var normalizedOperation = NormalizeOfficeKeyOperation(_settings.OfficeKeyOperation);
            if (string.Equals(normalizedOperation, "delete", StringComparison.OrdinalIgnoreCase))
            {
                return string.Format(
                    "当前操作：删除密钥文件首行。点击后只删除密钥源文件中的第一条有效密钥，不会写入桌面文件，不会复制到剪贴板，也不会打开 Office。\r\n剩余密钥：{0} 个\r\n操作计数：{1}\r\nOffice 当前状态：{2}\r\n右击此按钮可切换复制/剪切/删除密钥文件首行，也可发送自定义按键串。",
                    GetAvailableOfficeKeyCount(),
                    BuildOfficeActionCountText(),
                    BuildOfficeActivationStatusSummary());
            }

            return string.Format(
                "当前操作：{0}。点击后会把密钥写入目标文件、复制到剪贴板，并打开 Office 软件。\r\n剩余密钥：{1} 个\r\n操作计数：{2}\r\nOffice 当前状态：{3}\r\n右击此按钮可切换复制/剪切/删除密钥文件首行，也可发送自定义按键串。",
                MapOfficeActionToDisplay(normalizedOperation),
                GetAvailableOfficeKeyCount(),
                BuildOfficeActionCountText(),
                BuildOfficeActivationStatusSummary());
        }

        private string BuildOfficeActionCountText()
        {
            return string.Format("{0}，余{1}码", BuildCurrentOfficeActionCountText(), GetAvailableOfficeKeyCount());
        }

        private string BuildCurrentOfficeActionCountText()
        {
            var normalizedOperation = NormalizeOfficeKeyOperation(_settings.OfficeKeyOperation);
            var currentCount = 0;
            if (string.Equals(normalizedOperation, "cut", StringComparison.OrdinalIgnoreCase))
            {
                currentCount = Math.Max(0, _officeCutCount);
            }
            else if (string.Equals(normalizedOperation, "delete", StringComparison.OrdinalIgnoreCase))
            {
                currentCount = Math.Max(0, _officeDeleteCount);
            }
            else
            {
                currentCount = Math.Max(0, _officeCopyCount);
            }

            return string.Format("{0}{1}次", GetOfficeActionShortText(normalizedOperation), currentCount);
        }

        private string BuildOfficeActivationStatusSummary()
        {
            return string.IsNullOrWhiteSpace(_lastOfficeActivationStatus) ? "正在检测" : _lastOfficeActivationStatus;
        }

        private string BuildWifiActionSubtitle()
        {
            bool connected;
            string connectedSsid;
            if (TryGetCurrentWifiConnection(out connected, out connectedSsid) && connected)
            {
                return "已连 " + ShortenDisplayText(connectedSsid, 14);
            }

            return _settings.WifiProfiles.Count == 0 ? "未设置WiFi" : "未连接,点此联网";
        }

        private string BuildWifiActionTooltipText()
        {
            bool connected;
            string connectedSsid;
            var status = TryGetCurrentWifiConnection(out connected, out connectedSsid) && connected
                ? "当前已连接：" + connectedSsid
                : (_settings.WifiProfiles.Count == 0 ? "当前未设置 WiFi" : "当前未连接 WLAN");
            return status + "。左击会优先从设置里的第一个 WiFi 开始按顺序尝试连接；右击可直接选择要连接的 WiFi。";
        }

        private string BuildBatteryActionSubtitle()
        {
            if (!_batteryPresent || _batteryHealthPercent < 0)
            {
                return "无电池 / 未接入";
            }

            return string.Format("有电池 / 健康{0:0.#}%", _batteryHealthPercent);
        }

        private void UpdateBatteryActionButtonPresentation()
        {
            if (_batteryActionButton == null || _batteryActionButton.IsDisposed)
            {
                UpdateCompactModeBatteryButtons();
                return;
            }

            var info = _batteryActionButton.Tag as ButtonStateInfo;
            if (info == null)
            {
                return;
            }

            info.BaseTitle = "4. 电池检查";
            info.Subtitle = BuildBatteryActionSubtitle();
            _toolTip.SetToolTip(_batteryActionButton, BuildCompactBatteryButtonTooltipText());
            ApplyButtonVisualState(_batteryActionButton);
            UpdateCompactModeBatteryButtons();
        }

        private void UpdateCompactModeOfficeButtons()
        {
            UpdateCompactModeOfficeButtons(_zenModeWindow);
            UpdateCompactModeOfficeButtons(_speedrunModeWindow);
        }

        private void UpdateCompactModeBatteryButtons()
        {
            UpdateCompactModeBatteryButtons(_zenModeWindow);
            UpdateCompactModeBatteryButtons(_speedrunModeWindow);
        }

        private void UpdateCompactModeWifiButtons()
        {
            UpdateCompactModeWifiButtons(_zenModeWindow);
            UpdateCompactModeWifiButtons(_speedrunModeWindow);
        }

        private void UpdateCompactModeOfficeButtons(Form window)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            foreach (var button in window.Controls.OfType<Button>())
            {
                if (!string.Equals(button.Name, "__compactOfficeButton", StringComparison.Ordinal))
                {
                    continue;
                }

                var info = button.Tag as ButtonStateInfo;
                if (info != null)
                {
                    info.BaseTitle = GetCompactOfficeButtonText();
                    ApplyButtonVisualState(button);
                }
                else
                {
                    button.Text = GetCompactOfficeButtonText();
                }

                _toolTip.SetToolTip(button, BuildOfficeActionTooltipText());
            }
        }

        private void UpdateCompactModeBatteryButtons(Form window)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            foreach (var button in window.Controls.OfType<Button>())
            {
                if (!string.Equals(button.Name, "__compactBatteryButton", StringComparison.Ordinal))
                {
                    continue;
                }

                var info = button.Tag as ButtonStateInfo;
                if (info != null)
                {
                    info.BaseTitle = GetCompactBatteryButtonText();
                    ApplyButtonVisualState(button);
                }
                else
                {
                    button.Text = GetCompactBatteryButtonText();
                }

                _toolTip.SetToolTip(button, BuildCompactBatteryButtonTooltipText());
            }
        }

        private void UpdateCompactModeWifiButtons(Form window)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            foreach (var button in window.Controls.OfType<Button>())
            {
                if (!string.Equals(button.Name, "__compactWifiButton", StringComparison.Ordinal))
                {
                    continue;
                }

                var info = button.Tag as ButtonStateInfo;
                if (info != null)
                {
                    info.BaseTitle = GetCompactWifiButtonText();
                    ApplyButtonVisualState(button);
                }
                else
                {
                    button.Text = GetCompactWifiButtonText();
                }

                _toolTip.SetToolTip(button, BuildWifiActionTooltipText());
            }
        }

        private void SyncCompactModeButtonStates()
        {
            SyncCompactModeButtonStates(_zenModeWindow);
            SyncCompactModeButtonStates(_speedrunModeWindow);
        }

        private void SyncCompactModeButtonStates(Form window)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            foreach (var button in window.Controls.OfType<Button>())
            {
                var info = button.Tag as ButtonStateInfo;
                if (info == null || info.LinkedButton == null || info.LinkedButton.IsDisposed)
                {
                    continue;
                }

                var linkedInfo = info.LinkedButton.Tag as ButtonStateInfo;
                if (linkedInfo == null)
                {
                    continue;
                }

                var shouldUpdate = !string.Equals(info.State, linkedInfo.State, StringComparison.Ordinal)
                    || (string.Equals(button.Name, "__compactWifiButton", StringComparison.Ordinal) && !string.Equals(info.BaseTitle, GetCompactWifiButtonText(), StringComparison.Ordinal))
                    || (string.Equals(button.Name, "__compactOfficeButton", StringComparison.Ordinal) && !string.Equals(info.BaseTitle, GetCompactOfficeButtonText(), StringComparison.Ordinal))
                    || (string.Equals(button.Name, "__compactBatteryButton", StringComparison.Ordinal) && !string.Equals(info.BaseTitle, GetCompactBatteryButtonText(), StringComparison.Ordinal));
                if (!shouldUpdate)
                {
                    continue;
                }

                info.State = linkedInfo.State;
                if (string.Equals(button.Name, "__compactWifiButton", StringComparison.Ordinal))
                {
                    info.BaseTitle = GetCompactWifiButtonText();
                }
                else if (string.Equals(button.Name, "__compactOfficeButton", StringComparison.Ordinal))
                {
                    info.BaseTitle = GetCompactOfficeButtonText();
                }
                else if (string.Equals(button.Name, "__compactBatteryButton", StringComparison.Ordinal))
                {
                    info.BaseTitle = GetCompactBatteryButtonText();
                }
                ApplyButtonVisualState(button);
            }
        }

        private void SyncModeWindowsTopMost()
        {
            foreach (var window in new[] { _zenModeWindow, _speedrunModeWindow })
            {
                if (window == null || window.IsDisposed)
                {
                    continue;
                }

                window.TopMost = TopMost;
            }
        }

        private bool HasCompactModeWindowOpen()
        {
            return (_zenModeWindow != null && !_zenModeWindow.IsDisposed && _zenModeWindow.Visible)
                || (_speedrunModeWindow != null && !_speedrunModeWindow.IsDisposed && _speedrunModeWindow.Visible);
        }

        private void HideMainToTray(string modeName)
        {
            if (_modeNotifyIcon != null)
            {
                _modeNotifyIcon.Text = string.Format("{0} - {1}", GetConfiguredAppTitle(), modeName);
                _modeNotifyIcon.Visible = true;
            }

            if (Visible)
            {
                ShowInTaskbar = false;
                Hide();
            }
        }

        private void RestoreMainFromTray()
        {
            if (IsDisposed)
            {
                return;
            }

            if (_modeNotifyIcon != null)
            {
                _modeNotifyIcon.Visible = false;
                _modeNotifyIcon.Text = GetConfiguredAppTitle();
            }

            if (!Visible)
            {
                Show();
            }

            if (Opacity <= 0d)
            {
                Opacity = 1d;
            }

            ShowInTaskbar = true;
            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
            }

            Activate();
        }

        private void ApplyModeWindowTheme(Form window)
        {
            if (window == null || window.IsDisposed)
            {
                return;
            }

            var dark = IsDarkTheme();
            var formBack = dark ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);
            var cardBack = dark ? Color.FromArgb(30, 41, 59) : Color.White;
            var primaryText = dark ? Color.FromArgb(226, 232, 240) : Color.FromArgb(15, 23, 42);
            var borderColor = dark ? Color.FromArgb(71, 85, 105) : Color.FromArgb(203, 213, 225);

            window.BackColor = formBack;

            foreach (var button in window.Controls.OfType<Button>())
            {
                var info = button.Tag as ButtonStateInfo;
                button.FlatAppearance.BorderColor = borderColor;
                if (string.Equals(button.Name, "__compactThemeButton", StringComparison.Ordinal))
                {
                    button.Text = GetCompactThemeButtonText();
                    button.BackColor = formBack;
                    button.ForeColor = dark ? Color.FromArgb(147, 197, 253) : Color.FromArgb(29, 78, 216);
                    button.FlatAppearance.BorderSize = 0;
                    button.FlatAppearance.MouseOverBackColor = dark ? Color.FromArgb(30, 41, 59) : Color.FromArgb(239, 246, 255);
                    button.FlatAppearance.MouseDownBackColor = dark ? Color.FromArgb(30, 41, 59) : Color.FromArgb(219, 234, 254);
                }
                else if (IsCompactBlueLinkButtonName(button.Name))
                {
                    button.BackColor = formBack;
                    button.ForeColor = dark ? Color.FromArgb(147, 197, 253) : Color.FromArgb(29, 78, 216);
                    button.FlatAppearance.BorderSize = 0;
                    button.FlatAppearance.MouseOverBackColor = dark ? Color.FromArgb(30, 41, 59) : Color.FromArgb(239, 246, 255);
                    button.FlatAppearance.MouseDownBackColor = dark ? Color.FromArgb(30, 41, 59) : Color.FromArgb(219, 234, 254);
                }
                else if (string.Equals(button.Name, "__compactOfficeMenuButton", StringComparison.Ordinal))
                {
                    button.BackColor = dark ? Color.FromArgb(51, 65, 85) : Color.FromArgb(248, 250, 252);
                    button.ForeColor = dark ? Color.FromArgb(226, 232, 240) : Color.FromArgb(51, 65, 85);
                }
                else if (info != null && info.LinkedButton != null)
                {
                    ApplyButtonVisualState(button);
                }
                else
                {
                    button.BackColor = cardBack;
                    button.ForeColor = primaryText;
                }
            }

            UpdateCompactModeOfficeButtons(window);
            UpdateCompactModeSummary(window);
            ApplyCompactModeSummaryTheme(window, cardBack, primaryText);
            foreach (var menuButton in window.Controls.OfType<Button>().Where(button => string.Equals(button.Name, "__compactOfficeMenuButton", StringComparison.Ordinal)))
            {
                _toolTip.SetToolTip(menuButton, string.Format("当前操作：{0}。点击尖括号可切换。", MapOfficeActionToDisplay(_settings.OfficeKeyOperation)));
            }
        }

        private void ToggleDisplayMode(DisplayMode mode)
        {
            if (mode == DisplayMode.Normal)
            {
                _currentDisplayMode = DisplayMode.Normal;
            }
            else
            {
                if (_currentDisplayMode == DisplayMode.Normal)
                {
                    _normalModeBounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
                    _normalModeWasMaximized = WindowState == FormWindowState.Maximized;
                }

                _currentDisplayMode = _currentDisplayMode == mode ? DisplayMode.Normal : mode;
            }

            ApplyDisplayMode();
            PerformResponsiveLayout();
            ApplyTheme();
        }

        private void ApplyDisplayMode()
        {
            var isNormal = _currentDisplayMode == DisplayMode.Normal;
            var isZen = _currentDisplayMode == DisplayMode.Zen;
            var isSpeedrun = _currentDisplayMode == DisplayMode.Speedrun;

            if (isNormal)
            {
                FormBorderStyle = FormBorderStyle.Sizable;
                MaximizeBox = true;
                MinimumSize = new Size(300, 400);
                MaximumSize = Size.Empty;
                if (_normalModeBounds != Rectangle.Empty)
                {
                    WindowState = FormWindowState.Normal;
                    Bounds = _normalModeBounds;
                }
                if (_normalModeWasMaximized)
                {
                    WindowState = FormWindowState.Maximized;
                }
            }
            else
            {
                if (_currentDisplayMode != DisplayMode.Normal && _normalModeBounds == Rectangle.Empty)
                {
                    _normalModeBounds = Bounds;
                }

                WindowState = FormWindowState.Normal;
                FormBorderStyle = FormBorderStyle.FixedSingle;
                MaximizeBox = false;
                MinimumSize = new Size(360, 420);
                MaximumSize = new Size(360, 420);
                ClientSize = new Size(360, 420);
            }

            if (_titleLabel != null)
            {
                _titleLabel.Visible = isNormal;
            }

            if (_subLabel != null)
            {
                _subLabel.Visible = isNormal;
            }

            if (_infoPanel != null)
            {
                _infoPanel.Visible = isNormal;
            }

            if (_statusPanel != null)
            {
                _statusPanel.Visible = isNormal;
            }

            if (_batteryPanel != null)
            {
                _batteryPanel.Visible = isNormal;
            }

            if (_noteTextBox != null)
            {
                _noteTextBox.Visible = isNormal;
            }

            if (_speakerTestButton != null)
            {
                _speakerTestButton.Visible = isNormal || isZen;
            }

            if (_settingsButton != null)
            {
                _settingsButton.Visible = isNormal;
            }

            if (_topMostButton != null)
            {
                _topMostButton.Visible = isNormal;
            }

            if (_zenModeButton != null)
            {
                _zenModeButton.Visible = isNormal;
            }

            if (_speedrunModeButton != null)
            {
                _speedrunModeButton.Visible = isNormal;
            }

            if (_exitButton != null)
            {
                _exitButton.Visible = !isNormal;
            }

            if (_wifiActionButton != null) _wifiActionButton.Visible = isNormal || isZen;
            if (_windowsActionButton != null) _windowsActionButton.Visible = isNormal || isZen || isSpeedrun;
            if (_configActionButton != null) _configActionButton.Visible = isNormal || isZen || isSpeedrun;
            if (_batteryActionButton != null) _batteryActionButton.Visible = isNormal || isZen || isSpeedrun;
            if (_cameraActionButton != null) _cameraActionButton.Visible = isNormal || isZen;
            if (_keyboardActionButton != null) _keyboardActionButton.Visible = isNormal || isZen;
            if (_officeActionButton != null) _officeActionButton.Visible = isNormal || isZen || isSpeedrun;
            if (_officeOperationDropButton != null) _officeOperationDropButton.Visible = _officeActionButton != null && _officeActionButton.Visible;
            if (_clearActionButton != null) _clearActionButton.Visible = isNormal || isZen;
            if (_powerActionButton != null) _powerActionButton.Visible = isNormal || isZen;

            if (_buttonsPanel != null)
            {
                _buttonsPanel.Visible = true;
                foreach (Control child in _buttonsPanel.Controls)
                {
                    if (child.Name == "__cardTitle" || child.Name == "__cardSubtitle")
                    {
                        child.Visible = isNormal;
                    }
                }
            }

            if (_subLabel != null)
            {
                _subLabel.Text = "交机前快速核验系统、硬件与常用功能。";
            }
        }

        private void ApplyHeaderButtonStyle(Button button, bool active, bool danger, Color cardBack, Color primaryText, Color borderColor)
        {
            if (button == null)
            {
                return;
            }

            if (danger)
            {
                button.BackColor = Color.FromArgb(254, 242, 242);
                button.ForeColor = Color.FromArgb(185, 28, 28);
                button.FlatAppearance.BorderColor = Color.FromArgb(252, 165, 165);
                return;
            }

            if (active)
            {
                button.BackColor = IsDarkTheme() ? Color.FromArgb(30, 64, 175) : Color.FromArgb(219, 234, 254);
                button.ForeColor = IsDarkTheme() ? Color.White : Color.FromArgb(29, 78, 216);
                button.FlatAppearance.BorderColor = IsDarkTheme() ? Color.FromArgb(96, 165, 250) : Color.FromArgb(147, 197, 253);
                return;
            }

            button.BackColor = cardBack;
            button.ForeColor = primaryText;
            button.FlatAppearance.BorderColor = borderColor;
        }

        private void PerformResponsiveLayout()
        {
            try
            {
                var margin = 24;
                var headerTop = 18;
                var viewportWidth = Math.Max(280, ClientSize.Width - (VerticalScroll.Visible ? SystemInformation.VerticalScrollBarWidth : 0));
                var contentWidth = Math.Max(252, viewportWidth - margin * 2);
                var gap = 20;
                var narrow = contentWidth < 760;
                var compactHeader = viewportWidth < 450;
                var headerBottom = headerTop;

                if (compactHeader)
                {
                    var x = margin;
                    var y = headerTop;
                    var rowHeight = 0;
                    foreach (var control in new Control[] { _themeButton, _settingsButton, _topMostButton })
                    {
                        if (control == null || !control.Visible)
                        {
                            continue;
                        }

                        if (x + control.Width > viewportWidth - margin && x > margin)
                        {
                            x = margin;
                            y += rowHeight + 6;
                            rowHeight = 0;
                        }

                        control.Location = new Point(x, y);
                        x += control.Width + 8;
                        rowHeight = Math.Max(rowHeight, control.Height);
                        headerBottom = Math.Max(headerBottom, y + rowHeight);
                    }
                }
                else
                {
                    var headerCursorRight = viewportWidth - margin;
                    var headerButtons = new Control[] { _topMostButton, _settingsButton, _themeButton };
                    foreach (var control in headerButtons)
                    {
                        if (control == null || !control.Visible)
                        {
                            continue;
                        }

                        control.Location = new Point(headerCursorRight - control.Width, headerTop + 8);
                        headerCursorRight = control.Left - 8;
                    }

                    headerBottom = headerTop + 34;

                    if (_titleLabel != null && _titleLabel.Visible)
                    {
                        _titleLabel.Width = Math.Max(120, headerCursorRight - _titleLabel.Left - 18);
                    }

                    if (_subLabel != null && _subLabel.Visible)
                    {
                        _subLabel.Width = Math.Max(180, headerCursorRight - _subLabel.Left - 18);
                    }
                }

                if (_titleLabel != null && _titleLabel.Visible && compactHeader)
                {
                    _titleLabel.Location = new Point(margin, headerBottom + 10);
                    _titleLabel.Size = new Size(contentWidth, 34);
                    headerBottom = _titleLabel.Bottom;
                }

                if (_subLabel != null && _subLabel.Visible && compactHeader)
                {
                    _subLabel.Location = new Point(margin + 3, headerBottom + 2);
                    _subLabel.Size = new Size(contentWidth, 36);
                    headerBottom = _subLabel.Bottom;
                }

                var currentTop = (_titleLabel != null && _titleLabel.Visible && !compactHeader) ? 96 : headerBottom + 14;
                if (_infoPanel != null && _infoPanel.Visible)
                {
                    _infoPanel.Location = new Point(margin, currentTop);
                    _infoPanel.Size = new Size(narrow ? contentWidth : Math.Max(320, contentWidth - Math.Max(220, Math.Min(280, contentWidth / 3)) - gap), 340);
                    currentTop = _infoPanel.Bottom + 12;
                }

                if (_batteryPanel != null && _batteryPanel.Visible)
                {
                    if (_infoPanel != null && _infoPanel.Visible && !narrow)
                    {
                        _batteryPanel.Location = new Point(_infoPanel.Right + gap, 96);
                        _batteryPanel.Size = new Size(Math.Max(220, contentWidth - _infoPanel.Width - gap), 340);
                        currentTop = Math.Max(currentTop, _batteryPanel.Bottom + 12);
                    }
                    else
                    {
                        _batteryPanel.Location = new Point(margin, currentTop);
                        _batteryPanel.Size = new Size(contentWidth, 220);
                        currentTop = _batteryPanel.Bottom + 12;
                    }
                }

                if (_statusPanel != null && _statusPanel.Visible)
                {
                    _statusPanel.Location = new Point(margin, currentTop);
                    _statusPanel.Size = new Size(contentWidth, 142);
                    currentTop = _statusPanel.Bottom + 18;
                }

                if (_buttonsPanel != null && _buttonsPanel.Visible)
                {
                    _buttonsPanel.Location = new Point(margin, currentTop);
                    _buttonsPanel.Size = new Size(contentWidth, Math.Max(180, CalculateButtonsPanelHeight(contentWidth)));
                    LayoutActionButtons();
                    currentTop = _buttonsPanel.Bottom + 18;
                }

                var bottomUtilityButtons = new List<Button>();
                if (_speakerTestButton != null && _speakerTestButton.Visible)
                {
                    bottomUtilityButtons.Add(_speakerTestButton);
                }

                if (_zenModeButton != null && _zenModeButton.Visible)
                {
                    bottomUtilityButtons.Add(_zenModeButton);
                }

                if (_speedrunModeButton != null && _speedrunModeButton.Visible)
                {
                    bottomUtilityButtons.Add(_speedrunModeButton);
                }

                if (_noteTextBox != null && _noteTextBox.Visible)
                {
                    _noteTextBox.Location = new Point(margin + 4, currentTop);
                    _noteTextBox.Size = new Size(contentWidth, contentWidth >= 520 ? 42 : 56);
                    currentTop = _noteTextBox.Bottom + 10;
                }

                if (bottomUtilityButtons.Count > 0)
                {
                    var buttonGap = 6;
                    var buttonWidth = bottomUtilityButtons.Max(button => button.Width);
                    var buttonHeight = bottomUtilityButtons.Max(button => button.Height);
                    var singleRowWidth = bottomUtilityButtons.Count * buttonWidth + Math.Max(0, bottomUtilityButtons.Count - 1) * buttonGap;

                    if (contentWidth >= singleRowWidth)
                    {
                        var startX = margin + contentWidth - singleRowWidth;
                        for (var index = 0; index < bottomUtilityButtons.Count; index++)
                        {
                            bottomUtilityButtons[index].Location = new Point(startX + index * (buttonWidth + buttonGap), currentTop);
                        }

                        currentTop += buttonHeight + 18;
                    }
                    else if (contentWidth >= buttonWidth * 2 + buttonGap)
                    {
                        const int columns = 2;
                        var rows = (bottomUtilityButtons.Count + columns - 1) / columns;
                        var totalWidth = columns * buttonWidth + buttonGap;
                        var startX = margin + contentWidth - totalWidth;
                        for (var index = 0; index < bottomUtilityButtons.Count; index++)
                        {
                            var row = index / columns;
                            var column = index % columns;
                            bottomUtilityButtons[index].Location = new Point(
                                startX + column * (buttonWidth + buttonGap),
                                currentTop + row * (buttonHeight + buttonGap));
                        }

                        currentTop += rows * (buttonHeight + buttonGap) - buttonGap + 18;
                    }
                    else
                    {
                        var startX = margin + contentWidth - buttonWidth;
                        for (var index = 0; index < bottomUtilityButtons.Count; index++)
                        {
                            bottomUtilityButtons[index].Location = new Point(startX, currentTop + index * (buttonHeight + buttonGap));
                        }

                        currentTop += bottomUtilityButtons.Count * (buttonHeight + buttonGap) - buttonGap + 18;
                    }
                }

                AutoScrollMinSize = new Size(0, currentTop);
            }
            catch
            {
            }
        }

        private int CalculateButtonsPanelHeight(int panelWidth)
        {
            var visibleButtons = GetVisibleActionButtons();
            if (visibleButtons.Count == 0)
            {
                return 150;
            }

            var columns = panelWidth >= 760 ? 3 : panelWidth >= 500 ? 2 : 1;
            var rows = (visibleButtons.Count + columns - 1) / columns;
            return 48 + rows * 66 + Math.Max(0, rows - 1) * 12 + 18;
        }

        private void LayoutActionButtons()
        {
            if (_buttonsPanel == null || _buttonsPanel.IsDisposed)
            {
                return;
            }

            var actionButtons = GetVisibleActionButtons();

            if (actionButtons.Count == 0)
            {
                return;
            }

            const int leftPadding = 22;
            const int topPadding = 48;
            const int horizontalGap = 14;
            const int verticalGap = 12;
            const int buttonHeight = 66;
            var columns = _buttonsPanel.ClientSize.Width >= 760 ? 3 : _buttonsPanel.ClientSize.Width >= 500 ? 2 : 1;
            var availableWidth = Math.Max(220, _buttonsPanel.ClientSize.Width - leftPadding * 2);
            var buttonWidth = Math.Max(180, (availableWidth - horizontalGap * (columns - 1)) / columns);

            for (var index = 0; index < actionButtons.Count; index++)
            {
                var row = index / columns;
                var column = index % columns;
                var x = leftPadding + column * (buttonWidth + horizontalGap);
                var y = topPadding + row * (buttonHeight + verticalGap);
                var button = actionButtons[index];
                if (button == _officeActionButton && _officeOperationDropButton != null && _officeOperationDropButton.Visible)
                {
                    button.Location = new Point(x, y);
                    button.Size = new Size(buttonWidth, buttonHeight);
                    var dropWidth = 18;
                    var dropHeight = 18;
                    _officeOperationDropButton.Location = new Point(
                        button.Right - dropWidth - 8,
                        button.Top + Math.Max(4, (buttonHeight - dropHeight) / 2));
                    _officeOperationDropButton.Size = new Size(dropWidth, dropHeight);
                    _officeOperationDropButton.BringToFront();
                }
                else
                {
                    button.Location = new Point(x, y);
                    button.Size = new Size(buttonWidth, buttonHeight);
                }
            }
        }

        private List<Button> GetVisibleActionButtons()
        {
            return _buttonsPanel == null
                ? new List<Button>()
                : _buttonsPanel.Controls
                    .OfType<Button>()
                    .Where(button => button.Tag is ButtonStateInfo && button.Visible)
                    .OrderBy(button => GetActionButtonOrder(button.Tag as ButtonStateInfo))
                    .ToList();
        }

        private static int GetActionButtonOrder(ButtonStateInfo info)
        {
            if (info == null || string.IsNullOrWhiteSpace(info.BaseTitle))
            {
                return int.MaxValue;
            }

            var match = Regex.Match(info.BaseTitle, @"^\D*(\d+)");
            int order;
            return match.Success && int.TryParse(match.Groups[1].Value, out order) ? order : int.MaxValue;
        }

        private bool IsDarkTheme()
        {
            return string.Equals(_settings.ThemeMode, "dark", StringComparison.OrdinalIgnoreCase);
        }

        private void ApplyThemeButtonVisualState(Color cardBack, Color primaryText, Color borderColor)
        {
            if (_themeButton == null)
            {
                return;
            }

            var dark = IsDarkTheme();
            _themeButton.Text = dark ? "切换为亮主题" : "切换为暗主题";
            _themeButton.BackColor = cardBack;
            _themeButton.ForeColor = primaryText;
            _themeButton.FlatAppearance.BorderColor = borderColor;
            _toolTip.SetToolTip(
                _themeButton,
                dark
                    ? "当前为黑暗主题，点击切换为亮主题，并同步系统与 Office 主题"
                    : "当前为明亮主题，点击切换为黑暗主题，并同步系统与 Office 主题");
        }

        private void ApplyTopMostButtonVisualState()
        {
            if (_topMostButton == null)
            {
                return;
            }

            var dark = IsDarkTheme();
            _topMostButton.Text = _topMostButton.Checked ? "⬆ 已置顶" : "⇧ 未置顶";
            _topMostButton.FlatAppearance.BorderSize = _topMostButton.Checked ? 2 : 1;

            if (dark)
            {
                _topMostButton.BackColor = _topMostButton.Checked ? Color.FromArgb(51, 65, 85) : Color.FromArgb(30, 41, 59);
                _topMostButton.ForeColor = Color.FromArgb(226, 232, 240);
                _topMostButton.FlatAppearance.BorderColor = _topMostButton.Checked ? Color.FromArgb(148, 163, 184) : Color.FromArgb(71, 85, 105);
            }
            else
            {
                _topMostButton.BackColor = _topMostButton.Checked ? Color.FromArgb(255, 251, 235) : Color.White;
                _topMostButton.ForeColor = Color.FromArgb(15, 23, 42);
                _topMostButton.FlatAppearance.BorderColor = _topMostButton.Checked ? Color.FromArgb(245, 158, 11) : Color.FromArgb(203, 213, 225);
            }
        }

        private void ApplyTheme()
        {
            var dark = IsDarkTheme();
            var formBack = dark ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);
            var cardBack = dark ? Color.FromArgb(30, 41, 59) : Color.White;
            var primaryText = dark ? Color.FromArgb(226, 232, 240) : Color.FromArgb(15, 23, 42);
            var secondaryText = dark ? Color.FromArgb(148, 163, 184) : Color.FromArgb(71, 85, 105);
            var borderColor = dark ? Color.FromArgb(71, 85, 105) : Color.FromArgb(203, 213, 225);

            BackColor = formBack;
            ApplyThemeToControlTree(this, formBack, cardBack, primaryText, secondaryText);

            if (_settingsButton != null)
            {
                _settingsButton.BackColor = cardBack;
                _settingsButton.ForeColor = primaryText;
                _settingsButton.FlatAppearance.BorderColor = borderColor;
            }

            ApplyThemeButtonVisualState(cardBack, primaryText, borderColor);
            ApplyHeaderButtonStyle(_zenModeButton, _zenModeWindow != null && !_zenModeWindow.IsDisposed && _zenModeWindow.Visible, false, cardBack, primaryText, borderColor);
            ApplyHeaderButtonStyle(_speedrunModeButton, _speedrunModeWindow != null && !_speedrunModeWindow.IsDisposed && _speedrunModeWindow.Visible, false, cardBack, primaryText, borderColor);
            ApplyHeaderButtonStyle(_exitButton, false, false, cardBack, primaryText, borderColor);

            if (_topMostButton != null)
            {
                ApplyTopMostButtonVisualState();
            }

            if (_noteTextBox != null)
            {
                _noteTextBox.BackColor = formBack;
                _noteTextBox.ForeColor = secondaryText;
            }

            if (_speakerTestButton != null)
            {
                _speakerTestButton.BackColor = dark ? Color.FromArgb(30, 64, 175) : Color.FromArgb(219, 234, 254);
                _speakerTestButton.ForeColor = dark ? Color.White : Color.FromArgb(29, 78, 216);
                _speakerTestButton.FlatAppearance.BorderColor = dark ? Color.FromArgb(96, 165, 250) : Color.FromArgb(147, 197, 253);
            }

            if (_officeOperationDropButton != null)
            {
                _officeOperationDropButton.BackColor = dark ? Color.FromArgb(51, 65, 85) : Color.FromArgb(248, 250, 252);
                _officeOperationDropButton.ForeColor = dark ? Color.FromArgb(226, 232, 240) : Color.FromArgb(51, 65, 85);
                _officeOperationDropButton.FlatAppearance.BorderColor = dark ? Color.FromArgb(71, 85, 105) : Color.FromArgb(203, 213, 225);
            }

            foreach (var button in _trackedButtons)
            {
                ApplyButtonVisualState(button);
            }

            ApplyModeWindowTheme(_zenModeWindow);
            ApplyModeWindowTheme(_speedrunModeWindow);
        }

        private void ApplyThemeToControlTree(Control root, Color formBack, Color cardBack, Color primaryText, Color secondaryText)
        {
            foreach (Control control in root.Controls)
            {
                if (control == _settingsButton || control == _themeButton || control == _topMostButton || control == _speakerTestButton || control == _zenModeButton || control == _speedrunModeButton || control == _exitButton || control == _officeOperationDropButton)
                {
                    continue;
                }

                if (control is Panel)
                {
                    control.BackColor = cardBack;
                }
                else if (control is TextBox)
                {
                    var textBox = (TextBox)control;
                    textBox.BackColor = control.Parent == this ? formBack : cardBack;
                    textBox.ForeColor = control == _noteTextBox ? secondaryText : primaryText;
                }
                else if (control is Label)
                {
                    var label = (Label)control;
                    label.BackColor = control.Parent == this ? formBack : cardBack;
                    label.ForeColor = label.Font.Bold || label.Font.Size >= 12F ? primaryText : secondaryText;
                }
                else if (control is Button)
                {
                    var button = (Button)control;
                    button.FlatAppearance.BorderColor = darken(cardBack);
                }

                ApplyThemeToControlTree(control, formBack, cardBack, primaryText, secondaryText);
            }
        }

        private static Color darken(Color color)
        {
            return Color.FromArgb(Math.Max(0, color.R - 24), Math.Max(0, color.G - 24), Math.Max(0, color.B - 24));
        }

        private void SyncExternalThemes()
        {
            try
            {
                var dark = string.Equals(_settings.ThemeMode, "dark", StringComparison.OrdinalIgnoreCase);
                using (var personalize = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    if (personalize != null)
                    {
                        personalize.SetValue("AppsUseLightTheme", dark ? 0 : 1, RegistryValueKind.DWord);
                        personalize.SetValue("SystemUsesLightTheme", dark ? 0 : 1, RegistryValueKind.DWord);
                    }
                }

                using (var themesKey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes"))
                {
                    if (themesKey != null)
                    {
                        var windowsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
                        var themePath = Path.Combine(windowsDirectory, "Resources", "Themes", dark ? "dark.theme" : "aero.theme");
                        if (File.Exists(themePath))
                        {
                            themesKey.SetValue("CurrentTheme", themePath, RegistryValueKind.String);
                        }
                    }
                }

                foreach (var version in new[] { "16.0", "15.0", "14.0" })
                {
                    using (var officeThemeKey = Registry.CurrentUser.CreateSubKey(string.Format(@"Software\Microsoft\Office\{0}\Common", version)))
                    {
                        if (officeThemeKey != null)
                        {
                            officeThemeKey.SetValue("UI Theme", dark ? 4 : 5, RegistryValueKind.DWord);
                        }
                    }
                }

                IntPtr result;
                SendMessageTimeout(new IntPtr(0xffff), WmSettingchange, IntPtr.Zero, "ImmersiveColorSet", SmtoAbortifhung, 200, out result);
            }
            catch
            {
            }
        }

        private Panel CreateCardPanel(Point location, Size size, string title, string subtitle)
        {
            var panel = new Panel
            {
                Location = location,
                Size = size,
                BackColor = Color.White
            };

            var titleLabel = new Label
            {
                Text = title,
                Font = new Font("Microsoft YaHei UI", 12.5F, FontStyle.Bold, GraphicsUnit.Point),
                ForeColor = Color.FromArgb(15, 23, 42),
                Location = new Point(18, 14),
                Size = new Size(size.Width - 36, 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Name = "__cardTitle"
            };

            var subtitleLabel = new Label
            {
                Text = subtitle,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                ForeColor = Color.FromArgb(100, 116, 139),
                Location = new Point(18, 38),
                Size = new Size(size.Width - 36, 18),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Name = "__cardSubtitle"
            };

            panel.Controls.Add(titleLabel);
            panel.Controls.Add(subtitleLabel);
            return panel;
        }

        private TextBox AddInfoRow(Control parent, string label, int top, int valueHeight = 22)
        {
            var keyLabel = new Label
            {
                Text = label,
                ForeColor = Color.FromArgb(71, 85, 105),
                Location = new Point(18, top),
                Size = new Size(78, valueHeight),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };

            var valueLabel = new TextBox
            {
                Text = "读取中...",
                ForeColor = Color.FromArgb(15, 23, 42),
                Location = new Point(102, top),
                Size = new Size(parent.Width - 120, valueHeight),
                BackColor = parent.BackColor,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                ShortcutsEnabled = true,
                Multiline = valueHeight > 24,
                WordWrap = true,
                TabStop = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            parent.Controls.Add(keyLabel);
            parent.Controls.Add(valueLabel);
            return valueLabel;
        }

        private Button CreateActionButton(string text, string subtitle, Point location, EventHandler onClick, bool trackCompletion)
        {
            var button = new Button
            {
                Text = text + "\r\n" + subtitle,
                Location = location,
                Size = new Size(246, 66),
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(51, 65, 85),
                UseVisualStyleBackColor = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(14, 6, 6, 6)
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(71, 85, 105);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(30, 41, 59);
            button.Tag = new ButtonStateInfo
            {
                BaseTitle = text,
                Subtitle = subtitle,
                State = "default",
                TrackCompletion = trackCompletion
            };
            ApplyButtonVisualState(button);
            _toolTip.SetToolTip(button, text + "：" + subtitle);
            button.Click += (_, e) =>
            {
                onClick(button, e);
                if (trackCompletion)
                {
                    var currentInfo = button.Tag as ButtonStateInfo;
                    if (currentInfo != null && currentInfo.State == "default")
                    {
                        SetButtonState(button, "clicked");
                    }
                }
            };
            if (trackCompletion)
            {
                button.MouseDown += ActionButtonMouseDown;
                _trackedButtons.Add(button);
            }
            return button;
        }

        private Font GetActionButtonFont(float size)
        {
            Font font;
            if (_actionButtonFontCache.TryGetValue(size, out font))
            {
                return font;
            }

            font = new Font("Microsoft YaHei UI", size, FontStyle.Bold, GraphicsUnit.Point);
            _actionButtonFontCache[size] = font;
            return font;
        }

        private void ApplyActionButtonTypography(Button button)
        {
            if (button == null || button.IsDisposed)
            {
                return;
            }

            var availableWidth = Math.Max(80, button.ClientSize.Width - button.Padding.Horizontal - 8);
            var availableHeight = Math.Max(28, button.ClientSize.Height - button.Padding.Vertical - 6);
            var sizes = new[] { 10f, 9.5f, 9f, 8.5f, 8f, 7.5f };
            foreach (var size in sizes)
            {
                var font = GetActionButtonFont(size);
                var measured = TextRenderer.MeasureText(
                    button.Text ?? string.Empty,
                    font,
                    new Size(availableWidth, int.MaxValue),
                    TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
                if (measured.Height <= availableHeight)
                {
                    button.Font = font;
                    var scale = size / 10f;
                    var isOfficeButton = button == _officeActionButton;
                    var rightPadding = isOfficeButton && _officeOperationDropButton != null && _officeOperationDropButton.Visible
                        ? Math.Max(18, (int)Math.Round(20 * scale))
                        : Math.Max(4, (int)Math.Round(6 * scale));
                    button.Padding = new Padding(
                        isOfficeButton ? Math.Max(6, (int)Math.Round(8 * scale)) : Math.Max(8, (int)Math.Round(14 * scale)),
                        Math.Max(4, (int)Math.Round(6 * scale)),
                        rightPadding,
                        Math.Max(4, (int)Math.Round(6 * scale)));
                    return;
                }
            }

            button.Font = GetActionButtonFont(7.5f);
            button.Padding = button == _officeActionButton && _officeOperationDropButton != null && _officeOperationDropButton.Visible
                ? new Padding(6, 4, 18, 4)
                : new Padding(8, 4, 4, 4);
        }

        private void ActionButtonMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                return;
            }

            var button = sender as Button;
            if (button == null)
            {
                return;
            }

            var info = button.Tag as ButtonStateInfo;
            if (info == null || !info.TrackCompletion)
            {
                return;
            }

            var targetButton = info.LinkedButton != null && !info.LinkedButton.IsDisposed ? info.LinkedButton : button;
            var targetInfo = targetButton.Tag as ButtonStateInfo;
            if (targetInfo == null || !targetInfo.TrackCompletion)
            {
                return;
            }

            if (targetButton == _officeActionButton || string.Equals(button.Name, "__compactOfficeButton", StringComparison.Ordinal))
            {
                ShowOfficeOperationMenu(button);
                return;
            }

            if (targetButton == _windowsActionButton)
            {
                ShowWindowsActivationMenu(button);
                return;
            }

            if (targetButton == _wifiActionButton || string.Equals(button.Name, "__compactWifiButton", StringComparison.Ordinal))
            {
                ShowWifiSelectionMenu(button);
                return;
            }

            if (targetButton == _batteryActionButton || string.Equals(button.Name, "__compactBatteryButton", StringComparison.Ordinal))
            {
                if (targetInfo.State == "passed")
                {
                    LoadBatteryInfo();
                    UpdateBatteryActionButtonPresentation();
                    UpdateCompactModeSummaries();
                }
                else
                {
                    SetButtonState(targetButton, "passed");
                    ShowInfo("已手动将电池标记为通过。\r\n如需按真实健康度/容量自动判色，可在“设置 > 电池规则”里调整阈值；再次右击已为绿色的电池按钮，会恢复真实检测结果。", "电池");
                }
                return;
            }

            SetButtonState(targetButton, targetInfo.State == "passed" ? "clicked" : "passed");
        }

        private void SetButtonState(Button button, string state)
        {
            var info = button.Tag as ButtonStateInfo;
            if (info == null)
            {
                return;
            }

            if (string.Equals(info.State, state, StringComparison.Ordinal))
            {
                return;
            }

            info.State = state;
            ApplyButtonVisualState(button);
            SyncCompactModeButtonStates();
            CheckCompletionState();
        }

        private void ApplyButtonVisualState(Button button)
        {
            var info = button.Tag as ButtonStateInfo;
            if (info == null)
            {
                return;
            }

            var compactModeStyle = info.CompactModeStyle;
            var keepCompactBaseTitle = compactModeStyle && string.Equals(button.Name, "__compactOfficeButton", StringComparison.Ordinal);
            if (info.State == "passed")
            {
                button.Text = compactModeStyle
                    ? (keepCompactBaseTitle ? info.BaseTitle : "✔ " + info.BaseTitle)
                    : "✔  " + info.BaseTitle + "\r\n" + info.Subtitle;
                button.BackColor = Color.FromArgb(232, 249, 238);
                button.ForeColor = Color.FromArgb(21, 128, 61);
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.BorderColor = Color.FromArgb(134, 239, 172);
                button.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 252, 231);
                button.FlatAppearance.MouseDownBackColor = Color.FromArgb(220, 252, 231);
            }
            else if (info.State == "failed")
            {
                button.Text = compactModeStyle
                    ? (keepCompactBaseTitle ? info.BaseTitle : "✖ " + info.BaseTitle)
                    : "✖  " + info.BaseTitle + "\r\n" + info.Subtitle;
                button.BackColor = Color.FromArgb(254, 242, 242);
                button.ForeColor = Color.FromArgb(185, 28, 28);
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.BorderColor = Color.FromArgb(252, 165, 165);
                button.FlatAppearance.MouseOverBackColor = Color.FromArgb(254, 226, 226);
                button.FlatAppearance.MouseDownBackColor = Color.FromArgb(254, 226, 226);
            }
            else if (info.State == "clicked")
            {
                button.Text = compactModeStyle ? info.BaseTitle : info.BaseTitle + "\r\n" + info.Subtitle;
                button.BackColor = Color.FromArgb(255, 237, 213);
                button.ForeColor = Color.FromArgb(194, 65, 12);
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.BorderColor = Color.FromArgb(253, 186, 116);
                button.FlatAppearance.MouseOverBackColor = Color.FromArgb(254, 215, 170);
                button.FlatAppearance.MouseDownBackColor = Color.FromArgb(254, 215, 170);
            }
            else
            {
                button.Text = compactModeStyle ? info.BaseTitle : info.BaseTitle + "\r\n" + info.Subtitle;
                button.BackColor = Color.FromArgb(71, 85, 105);
                button.ForeColor = Color.White;
                button.FlatAppearance.BorderSize = 0;
                button.FlatAppearance.MouseOverBackColor = Color.FromArgb(100, 116, 139);
                button.FlatAppearance.MouseDownBackColor = Color.FromArgb(51, 65, 85);
            }

            if (compactModeStyle)
            {
                var hasSecondaryLine = !string.IsNullOrWhiteSpace(info.BaseTitle) && info.BaseTitle.IndexOf('\n') >= 0;
                button.Padding = hasSecondaryLine ? new Padding(2) : Padding.Empty;
                if (string.Equals(button.Name, "__compactOfficeButton", StringComparison.Ordinal))
                {
                    button.Font = new Font("Microsoft YaHei UI", 6.2F, FontStyle.Bold, GraphicsUnit.Point);
                }
                else
                {
                    button.Font = hasSecondaryLine
                        ? new Font("Microsoft YaHei UI", 7.4F, FontStyle.Bold, GraphicsUnit.Point)
                        : (info.BaseTitle != null && info.BaseTitle.Length >= 10
                            ? new Font("Microsoft YaHei UI", 8F, FontStyle.Bold, GraphicsUnit.Point)
                            : new Font("Microsoft YaHei UI", 8.5F, FontStyle.Bold, GraphicsUnit.Point));
                }
            }
            else
            {
                ApplyActionButtonTypography(button);
            }
        }

        private void CheckCompletionState()
        {
            if (_completionShown)
            {
                return;
            }

            if (_trackedButtons.Count == 0)
            {
                return;
            }

            if (_trackedButtons.All(button =>
            {
                var info = button.Tag as ButtonStateInfo;
                return info != null && info.State == "passed";
            }))
            {
                _completionShown = true;
                ShowCompletionShutdownDialog();
            }
        }

        private void ShowCompletionShutdownDialog()
        {
            using (var dialog = new Form())
            {
                var remainingSeconds = 30;
                var timer = new System.Windows.Forms.Timer();
                dialog.Text = "检测完成";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.ClientSize = new Size(420, 210);
                dialog.BackColor = Color.FromArgb(240, 253, 244);
                dialog.Font = Font;

                var titleLabel = new Label
                {
                    Text = "恭喜，检测全通过",
                    Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold, GraphicsUnit.Point),
                    ForeColor = Color.FromArgb(21, 128, 61),
                    Location = new Point(28, 24),
                    Size = new Size(260, 28)
                };

                var bodyLabel = new Label
                {
                    Text = "所有检测按钮都已标记为通过，准备自动关机。",
                    ForeColor = Color.FromArgb(22, 101, 52),
                    Location = new Point(30, 68),
                    Size = new Size(320, 22)
                };

                var countdownLabel = new Label
                {
                    Text = "30 秒后自动关机",
                    Font = new Font("Microsoft YaHei UI", 13F, FontStyle.Bold, GraphicsUnit.Point),
                    ForeColor = Color.FromArgb(22, 101, 52),
                    Location = new Point(30, 102),
                    Size = new Size(220, 28)
                };

                var cancelButton = new Button
                {
                    Text = "取消关机",
                    Location = new Point(220, 150),
                    Size = new Size(100, 34)
                };

                var nowButton = new Button
                {
                    Text = "立即关机",
                    Location = new Point(100, 150),
                    Size = new Size(100, 34)
                };

                timer.Interval = 1000;
                timer.Tick += (_, __) =>
                {
                    remainingSeconds--;
                    countdownLabel.Text = string.Format("{0} 秒后自动关机", remainingSeconds);
                    if (remainingSeconds <= 0)
                    {
                        timer.Stop();
                        TryStartTarget("shutdown.exe", "/s /t 0");
                        dialog.Close();
                    }
                };

                cancelButton.Click += (_, __) =>
                {
                    timer.Stop();
                    TryStartTarget("shutdown.exe", "/a");
                    dialog.Close();
                    _completionShown = false;
                };

                nowButton.Click += (_, __) =>
                {
                    timer.Stop();
                    TryStartTarget("shutdown.exe", "/s /t 0");
                    dialog.Close();
                };

                dialog.FormClosed += (_, __) => timer.Stop();
                dialog.Controls.Add(titleLabel);
                dialog.Controls.Add(bodyLabel);
                dialog.Controls.Add(countdownLabel);
                dialog.Controls.Add(nowButton);
                dialog.Controls.Add(cancelButton);
                timer.Start();
                dialog.ShowDialog(this);
            }
        }

        private void RefreshDashboard()
        {
            LoadSystemInfo();
            LoadBatteryInfo();
            LoadStatusInfo();
            SyncWindowsActivationButtonState();
            UpdateBatteryActionButtonPresentation();
            UpdateWifiActionButtonPresentation();
            UpdateOfficeActionButtonPresentation();
            UpdateCompactModeWifiButtons();
            SyncOfficeActivationButtonStateAsync();
            SyncWifiActionButtonState();
            UpdateCompactModeSummaries();
        }

        private void UpdateOfficeActionButtonPresentation()
        {
            if (_officeActionButton == null || _officeActionButton.IsDisposed)
            {
                UpdateCompactModeOfficeButtons();
                return;
            }

            var info = _officeActionButton.Tag as ButtonStateInfo;
            if (info == null)
            {
                return;
            }

            info.BaseTitle = string.Format("7. Office激活 余{0}码", GetAvailableOfficeKeyCount());
            info.Subtitle = BuildOfficeActionSubtitle();
            _toolTip.SetToolTip(_officeActionButton, BuildOfficeActionTooltipText());
            ApplyButtonVisualState(_officeActionButton);
            UpdateCompactModeOfficeButtons();
            UpdateOfficeOperationDropButtonPresentation();
        }

        private void SyncOfficeActivationButtonStateAsync()
        {
            if (_officeActivationCheckInProgress || IsDisposed)
            {
                return;
            }

            _officeActivationCheckInProgress = true;
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                var activated = false;
                string statusText;
                TryGetOfficeActivationState(out activated, out statusText);

                try
                {
                    BeginInvoke((MethodInvoker)delegate
                    {
                        _officeActivationCheckInProgress = false;
                        _lastOfficeActivationStatus = statusText;
                        ApplyOfficeActivationState(activated, statusText);
                    });
                }
                catch
                {
                    _officeActivationCheckInProgress = false;
                }
            });
        }

        private void ApplyOfficeActivationState(bool activated, string statusText)
        {
            _lastOfficeActivationStatus = string.IsNullOrWhiteSpace(statusText)
                ? (activated ? "Office 已激活" : "Office 未激活")
                : statusText;

            if (_officeActionButton != null && !_officeActionButton.IsDisposed)
            {
                SetButtonState(_officeActionButton, activated ? "passed" : "failed");
                UpdateOfficeActionButtonPresentation();
            }
            else
            {
                UpdateCompactModeOfficeButtons();
            }
        }

        private bool TryGetOfficeActivationState(out bool activated, out string statusText)
        {
            foreach (var scriptPath in BuildOfficeActivationScriptCandidates())
            {
                if (!File.Exists(scriptPath))
                {
                    continue;
                }

                var output = RunProcessForOutput(
                    "cscript.exe",
                    string.Format("//Nologo \"{0}\" /dstatus", scriptPath),
                    Path.GetDirectoryName(scriptPath) ?? GetApplicationDirectory());

                if (string.IsNullOrWhiteSpace(output))
                {
                    continue;
                }

                if (output.IndexOf("---LICENSED---", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    activated = true;
                    statusText = "Office 已激活";
                    return true;
                }

                if (output.IndexOf("LICENSE STATUS", StringComparison.OrdinalIgnoreCase) >= 0
                    || output.IndexOf("许可证状态", StringComparison.OrdinalIgnoreCase) >= 0
                    || output.IndexOf("ライセンス状態", StringComparison.OrdinalIgnoreCase) >= 0
                    || output.IndexOf("No installed product keys detected", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    activated = false;
                    statusText = ExtractOfficeActivationStatus(output);
                    return true;
                }
            }

            activated = false;
            statusText = "未检测到 Office 激活信息";
            return false;
        }

        private void SyncWindowsActivationButtonState()
        {
            bool activated;
            string statusText;
            TryGetWindowsActivationState(out activated, out statusText);
            ApplyWindowsActivationState(activated, statusText);
        }

        private void ApplyWindowsActivationState(bool activated, string statusText)
        {
            _lastWindowsActivationStatus = string.IsNullOrWhiteSpace(statusText)
                ? (activated ? "Windows 已激活" : "Windows 未激活")
                : statusText;

            if (_windowsActionButton == null || _windowsActionButton.IsDisposed)
            {
                return;
            }

            var currentInfo = _windowsActionButton.Tag as ButtonStateInfo;
            if (!activated
                && currentInfo != null
                && string.Equals(currentInfo.State, "passed", StringComparison.Ordinal)
                && IsUnknownWindowsActivationStatus(_lastWindowsActivationStatus))
            {
                UpdateWindowsActivationButtonPresentation();
                return;
            }

            SetButtonState(_windowsActionButton, activated ? "passed" : "failed");
            UpdateWindowsActivationButtonPresentation();
        }

        private void UpdateWindowsActivationButtonPresentation()
        {
            if (_windowsActionButton == null || _windowsActionButton.IsDisposed)
            {
                return;
            }

            var info = _windowsActionButton.Tag as ButtonStateInfo;
            if (info == null)
            {
                return;
            }

            info.BaseTitle = "2. Windows 激活";
            info.Subtitle = "激活页 + 校准时间";
            _toolTip.SetToolTip(
                _windowsActionButton,
                string.Format(
                    "当前系统状态：{0}\r\n左击：先尝试校准系统时间，再打开 Windows 激活页面。\r\n右击：显示管理员 CMD 激活菜单。",
                    BuildWindowsActivationStatusSummary()));
            ApplyButtonVisualState(_windowsActionButton);
        }

        private string BuildWindowsActivationStatusSummary()
        {
            var activation = string.IsNullOrWhiteSpace(_lastWindowsActivationStatus) ? "正在检测激活状态" : _lastWindowsActivationStatus;
            var timeSync = string.IsNullOrWhiteSpace(_lastWindowsTimeSyncStatus) ? "尚未校准时间" : _lastWindowsTimeSyncStatus;
            return activation + "；" + timeSync;
        }

        private static bool IsUnknownWindowsActivationStatus(string statusText)
        {
            var text = (statusText ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            return text.IndexOf("未读取到", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("读取失败", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("未知", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool TryGetWindowsActivationState(out bool activated, out string statusText)
        {
            activated = false;
            statusText = "Windows 未激活";
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT Name, Description, LicenseStatus, PartialProductKey FROM SoftwareLicensingProduct WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL"))
                {
                    var products = searcher.Get().Cast<ManagementObject>().ToList();
                    if (products.Count == 0)
                    {
                        statusText = "未读取到 Windows 激活信息";
                        return false;
                    }

                    var licensedProduct = products.FirstOrDefault(item => (int)Math.Round(ToDouble(item["LicenseStatus"])) == 1);
                    if (licensedProduct != null)
                    {
                        activated = true;
                        statusText = "Windows 已激活";
                        return true;
                    }

                    var first = products[0];
                    statusText = "Windows 未激活";
                    var licenseStatus = (int)Math.Round(ToDouble(first["LicenseStatus"]));
                    switch (licenseStatus)
                    {
                        case 0:
                            statusText = "Windows 未授权";
                            break;
                        case 2:
                            statusText = "Windows 处于宽限期";
                            break;
                        case 3:
                            statusText = "Windows 处于宽限期";
                            break;
                        case 4:
                            statusText = "Windows 非正版/通知状态";
                            break;
                        case 5:
                            statusText = "Windows 处于扩展宽限期";
                            break;
                        case 6:
                            statusText = "Windows 扩展宽限期";
                            break;
                    }
                }
            }
            catch
            {
                statusText = "Windows 激活状态读取失败";
            }

            return activated;
        }

        private string TryResyncWindowsTime()
        {
            var workingDirectory = GetApplicationDirectory();
            var attempts = new List<ProcessExecutionResult>();

            var firstAttempt = RunProcessForOutputDetailed("w32tm.exe", "/resync /force", workingDirectory, 22000);
            attempts.Add(firstAttempt);
            if (IsSuccessfulWindowsTimeSyncOutput(firstAttempt.CombinedOutput))
            {
                return "时间已校准";
            }

            var startService = RunProcessForOutputDetailed("sc.exe", "start w32time", workingDirectory, 12000);
            attempts.Add(startService);
            System.Threading.Thread.Sleep(1200);

            var retryForce = RunProcessForOutputDetailed("w32tm.exe", "/resync /force", workingDirectory, 22000);
            attempts.Add(retryForce);
            if (IsSuccessfulWindowsTimeSyncOutput(retryForce.CombinedOutput))
            {
                return "时间已校准";
            }

            var retryNormal = RunProcessForOutputDetailed("w32tm.exe", "/resync", workingDirectory, 22000);
            attempts.Add(retryNormal);
            if (IsSuccessfulWindowsTimeSyncOutput(retryNormal.CombinedOutput))
            {
                return "时间已校准";
            }

            var finalAttempt = attempts
                .LastOrDefault(item => item != null && (!string.IsNullOrWhiteSpace(item.CombinedOutput) || item.TimedOut || !item.Started))
                ?? retryNormal;
            var mergedOutput = string.Join(
                "\r\n",
                attempts
                    .Where(item => item != null)
                    .Select(item => item.CombinedOutput)
                    .Where(text => !string.IsNullOrWhiteSpace(text)));
            return BuildWindowsTimeSyncFailureMessage(finalAttempt, mergedOutput);
        }

        private static bool IsSuccessfulWindowsTimeSyncOutput(string output)
        {
            var text = (output ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return text.IndexOf("命令成功完成", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("The command completed successfully", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("已成功完成重新同步", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("正しく完了しました", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string BuildWindowsTimeSyncFailureMessage(ProcessExecutionResult attempt, string mergedOutput)
        {
            var output = string.IsNullOrWhiteSpace(mergedOutput) ? (attempt == null ? string.Empty : attempt.CombinedOutput) : mergedOutput;
            var normalized = (output ?? string.Empty).Trim();
            if (attempt != null)
            {
                if (!attempt.Started)
                {
                    return "时间校准失败：时间命令未启动";
                }

                if (attempt.TimedOut)
                {
                    return "时间校准失败：执行超时";
                }
            }

            if (string.IsNullOrWhiteSpace(normalized))
            {
                return "时间校准失败";
            }

            if (ContainsAny(normalized, "拒绝访问", "access is denied", "需要提升", "requires elevation", "requested operation requires elevation", "需要特权"))
            {
                return "时间校准失败：权限不足";
            }

            if (ContainsAny(normalized, "服务尚未启动", "service has not been started", "服务无法启动", "service cannot be started"))
            {
                return "时间校准失败：时间服务未启动";
            }

            if (ContainsAny(normalized, "没有可用的时间数据", "no time data was available", "时间源", "time source"))
            {
                return "时间校准失败：当前无可用时间源";
            }

            if (ContainsAny(normalized, "找不到文件", "not recognized", "无法将", "could not find", "系统找不到"))
            {
                return "时间校准失败：系统缺少时间校准命令";
            }

            if (ContainsAny(normalized, "rpc 服务器不可用", "rpc server is unavailable"))
            {
                return "时间校准失败：时间服务异常";
            }

            var firstLine = ExtractFirstMeaningfulLine(normalized);
            return string.IsNullOrWhiteSpace(firstLine)
                ? "时间校准失败"
                : "时间校准失败：" + firstLine;
        }

        private static bool ContainsAny(string text, params string[] patterns)
        {
            if (string.IsNullOrWhiteSpace(text) || patterns == null)
            {
                return false;
            }

            return patterns.Any(pattern => !string.IsNullOrWhiteSpace(pattern)
                && text.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static string ExtractFirstMeaningfulLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            foreach (var rawLine in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var line = rawLine.Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                return line;
            }

            return string.Empty;
        }

        private sealed class ProcessExecutionResult
        {
            public ProcessExecutionResult()
            {
                ExitCode = int.MinValue;
            }

            public bool Started { get; set; }
            public bool TimedOut { get; set; }
            public int ExitCode { get; set; }
            public string StandardOutput { get; set; }
            public string StandardError { get; set; }
            public string ExceptionMessage { get; set; }

            public string CombinedOutput
            {
                get
                {
                    return string.Join(
                            "\r\n",
                            new[] { StandardOutput, StandardError, ExceptionMessage }
                                .Where(part => !string.IsNullOrWhiteSpace(part)))
                        .Trim();
                }
            }
        }

        private IEnumerable<string> BuildOfficeActivationScriptCandidates()
        {
            var candidates = new List<string>();
            Action<string> addCandidate = candidate =>
            {
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    return;
                }

                if (!candidates.Any(existing => string.Equals(existing, candidate, StringComparison.OrdinalIgnoreCase)))
                {
                    candidates.Add(candidate);
                }
            };

            foreach (var programRoot in new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            }.Where(path => !string.IsNullOrWhiteSpace(path)))
            {
                foreach (var officeFolder in new[]
                {
                    @"Microsoft Office\Office16\OSPP.VBS",
                    @"Microsoft Office\Office15\OSPP.VBS",
                    @"Microsoft Office\Office14\OSPP.VBS",
                    @"Microsoft Office\root\Office16\OSPP.VBS",
                    @"Microsoft Office\root\Office15\OSPP.VBS",
                    @"Microsoft Office\root\Office14\OSPP.VBS"
                })
                {
                    addCandidate(Path.Combine(programRoot, officeFolder));
                }
            }

            return candidates;
        }

        private static string RunProcessForOutput(string fileName, string arguments, string workingDirectory)
        {
            return RunProcessForOutputDetailed(fileName, arguments, workingDirectory, 15000).CombinedOutput;
        }

        private static ProcessExecutionResult RunProcessForOutputDetailed(string fileName, string arguments, string workingDirectory, int timeoutMs)
        {
            var result = new ProcessExecutionResult();
            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();

            try
            {
                using (var process = new Process())
                {
                    process.StartInfo = new ProcessStartInfo
                    {
                        FileName = fileName,
                        Arguments = arguments ?? string.Empty,
                        WorkingDirectory = string.IsNullOrWhiteSpace(workingDirectory) ? Environment.CurrentDirectory : workingDirectory,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    process.OutputDataReceived += (_, eventArgs) =>
                    {
                        if (eventArgs.Data == null)
                        {
                            return;
                        }

                        lock (outputBuilder)
                        {
                            outputBuilder.AppendLine(eventArgs.Data);
                        }
                    };
                    process.ErrorDataReceived += (_, eventArgs) =>
                    {
                        if (eventArgs.Data == null)
                        {
                            return;
                        }

                        lock (errorBuilder)
                        {
                            errorBuilder.AppendLine(eventArgs.Data);
                        }
                    };

                    if (!process.Start())
                    {
                        result.ExceptionMessage = "进程未启动。";
                        return result;
                    }

                    result.Started = true;
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    if (!process.WaitForExit(timeoutMs))
                    {
                        result.TimedOut = true;
                        try
                        {
                            process.Kill();
                        }
                        catch
                        {
                        }
                    }

                    try
                    {
                        process.WaitForExit();
                    }
                    catch
                    {
                    }

                    result.ExitCode = process.HasExited ? process.ExitCode : -1;
                }
            }
            catch (Exception ex)
            {
                result.ExceptionMessage = ex.Message;
            }

            result.StandardOutput = outputBuilder.ToString().Trim();
            result.StandardError = errorBuilder.ToString().Trim();
            return result;
        }

        private static string ExtractOfficeActivationStatus(string output)
        {
            var normalized = (output ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return "Office 激活状态未知";
            }

            if (normalized.IndexOf("No installed product keys detected", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "未检测到 Office 产品密钥";
            }

            foreach (var status in new[]
            {
                "---UNLICENSED---",
                "---NOTIFICATIONS---",
                "---OOB_GRACE---",
                "---GRACE---",
                "---RFM---"
            })
            {
                if (normalized.IndexOf(status, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "Office 未激活";
                }
            }

            var match = Regex.Match(normalized, @"LICENSE STATUS:\s*(.+)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (match.Success)
            {
                return "Office 状态：" + match.Groups[1].Value.Trim();
            }

            return "Office 未激活";
        }

        private static void CopyTextToClipboard(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new InvalidOperationException("没有可复制到剪贴板的密钥内容。");
            }

            Exception lastError = null;
            for (var i = 0; i < 5; i++)
            {
                try
                {
                    var dataObject = new DataObject();
                    dataObject.SetText(text);
                    Clipboard.SetDataObject(dataObject, true);
                    var copied = Clipboard.GetText();
                    if (string.Equals(copied, text, StringComparison.Ordinal))
                    {
                        return;
                    }
                }
                catch (Exception ex)
                {
                    lastError = ex;
                }

                System.Threading.Thread.Sleep(80);
            }

            throw new InvalidOperationException(
                lastError == null
                    ? "复制到剪贴板失败。"
                    : string.Format("复制到剪贴板失败：{0}", lastError.Message));
        }

        private void LoadSystemInfo()
        {
            try
            {
                using (var osSearcher = new ManagementObjectSearcher("SELECT Caption, Version, OSArchitecture, BuildNumber FROM Win32_OperatingSystem"))
                using (var csSearcher = new ManagementObjectSearcher("SELECT Manufacturer, Model, TotalPhysicalMemory FROM Win32_ComputerSystem"))
                using (var cpuSearcher = new ManagementObjectSearcher("SELECT Name FROM Win32_Processor"))
                using (var memorySearcher = new ManagementObjectSearcher("SELECT Manufacturer, Capacity, DeviceLocator, BankLabel, PartNumber FROM Win32_PhysicalMemory"))
                using (var diskSearcher = new ManagementObjectSearcher("SELECT Model, Size, MediaType, InterfaceType, SerialNumber, PNPDeviceID FROM Win32_DiskDrive"))
                using (var logicalDiskSearcher = new ManagementObjectSearcher("SELECT Size, FreeSpace, DriveType FROM Win32_LogicalDisk WHERE DriveType = 3"))
                {
                    var os = osSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                    var computer = csSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                    var cpu = cpuSearcher.Get().Cast<ManagementObject>().FirstOrDefault();
                    var memories = memorySearcher.Get().Cast<ManagementObject>().ToList();
                    var disks = diskSearcher.Get().Cast<ManagementObject>().ToList();
                    var logicalDisks = logicalDiskSearcher.Get().Cast<ManagementObject>().ToList();

                    _osLabel.Text = os == null
                        ? "无数据"
                        : BuildWindowsSummary(
                            SafeToString(os["Caption"]),
                            SafeToString(os["OSArchitecture"]),
                            SafeToString(os["BuildNumber"]));

                    if (computer == null)
                    {
                        _modelLabel.Text = "无数据";
                        _memoryLabel.Text = "无数据";
                    }
                    else
                    {
                        _computerManufacturer = SafeToString(computer["Manufacturer"]);
                        _computerModel = SafeToString(computer["Model"]);
                        _modelLabel.Text = string.Format("{0} {1}", _computerManufacturer, _computerModel).Trim();
                        _memoryLabel.Text = BuildMemorySummary(
                            memories,
                            ToDouble(computer["TotalPhysicalMemory"]),
                            GetVisiblePhysicalMemoryBytes());
                    }

                    _cpuLabel.Text = cpu == null ? "无数据" : SafeToString(cpu["Name"]);

                    if (disks.Count == 0)
                    {
                        _diskLabel.Text = "无数据";
                    }
                    else
                    {
                        _diskLabel.Text = BuildDiskSummary(disks, logicalDisks);
                    }
                }
            }
            catch
            {
                _osLabel.Text = "读取失败";
                _modelLabel.Text = "读取失败";
                _cpuLabel.Text = "读取失败";
                _memoryLabel.Text = "读取失败";
                _diskLabel.Text = "读取失败";
            }
        }

        private void LoadBatteryInfo()
        {
            try
            {
                bool hasBatteryDevice;
                var batteryDeviceDetected = TryDetectBatteryDeviceFast(out hasBatteryDevice);

                if (TryLoadBatteryInfoFromBatteryInfoView())
                {
                    return;
                }

                var reportPath = Path.Combine(Path.GetTempPath(), "pc-check-battery-report.html");
                try
                {
                    RunCommand("powercfg", string.Format("/batteryreport /output \"{0}\"", reportPath));
                }
                catch
                {
                }

                if (TryLoadBatteryInfoFromBatteryReport(reportPath))
                {
                    return;
                }

                if (batteryDeviceDetected && !hasBatteryDevice)
                {
                    SetBatteryUnavailable("未检测到电池设备");
                    return;
                }

                SetBatteryUnavailable("未读取到有效电池信息");
            }
            catch
            {
                SetBatteryUnavailable("电池信息读取失败");
            }
        }

        private bool TryLoadBatteryInfoFromBatteryInfoView()
        {
            var toolPath = FindBatteryToolPath();
            if (string.IsNullOrWhiteSpace(toolPath))
            {
                return false;
            }

            var exportPath = Path.Combine(Path.GetTempPath(), "batteryinfoview.txt");
            RunCommand(toolPath, string.Format("/stab \"{0}\"", exportPath));
            if (!File.Exists(exportPath))
            {
                return false;
            }

            var map = LoadBatteryInfoViewMap(exportPath);
            if (map.Count == 0)
            {
                return false;
            }

            var fullChargeCapacity = ParseBatteryNumber(GetDictionaryValue(map, "电池满电容量", "Full Charged Capacity", "Full Charge Capacity", "完全充電時の容量", "満充電容量"));
            var designCapacity = ParseBatteryNumber(GetDictionaryValue(map, "电池设计容量", "Designed Capacity", "Design Capacity", "設計容量", "バッテリー設計容量"));

            if (fullChargeCapacity <= 0 || designCapacity <= 0)
            {
                return false;
            }

            var powerState = GetDictionaryValue(map, "电源状态", "Power State", "電源状態", "電源の状態");
            var status = string.IsNullOrWhiteSpace(powerState)
                ? "来自 BatteryInfoView"
                : string.Format("来自 BatteryInfoView | {0}", powerState);

            ApplyBatteryMetrics(
                fullChargeCapacity,
                designCapacity,
                status,
                FormatBatteryToolCapacity,
                FormatBatteryToolCapacity);
            return true;
        }

        private bool TryLoadBatteryInfoFromBatteryReport(string reportPath)
        {
            if (!File.Exists(reportPath))
            {
                return false;
            }

            foreach (var encoding in GetBatteryTextEncodings())
            {
                try
                {
                    var html = File.ReadAllText(reportPath, encoding);
                    var designCapacity = ExtractBatteryValue(html, "DESIGN CAPACITY", "設計容量", "电池设计容量");
                    var fullChargeCapacity = ExtractBatteryValue(html, "FULL CHARGE CAPACITY", "FULL CHARGED CAPACITY", "完全充電時の容量", "満充電容量", "电池满电容量");
                    if (designCapacity <= 0 || fullChargeCapacity <= 0)
                    {
                        continue;
                    }

                    var cycleCount = ExtractBatteryText(html, "CYCLE COUNT", "サイクル カウント", "循环次数");
                    var status = string.IsNullOrWhiteSpace(cycleCount)
                        ? "来自 Windows 电池报告"
                        : string.Format("来自 Windows 电池报告 | 循环次数 {0}", cycleCount);

                    ApplyBatteryMetrics(
                        fullChargeCapacity,
                        designCapacity,
                        status,
                        FormatMilliwattHour,
                        FormatMilliwattHour);
                    return true;
                }
                catch
                {
                }
            }

            return false;
        }

        private string FindBatteryToolPath()
        {
            return BuildPortableCandidates(
                string.Empty,
                "BatteryInfoView.exe",
                @"XiaoGongJu\图吧工具箱202507\BatteryInfoView\BatteryInfoView.exe").FirstOrDefault(File.Exists);
        }

        private static string GetDictionaryValue(Dictionary<string, string> map, params string[] keys)
        {
            if (map == null || map.Count == 0 || keys == null || keys.Length == 0)
            {
                return string.Empty;
            }

            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                string value;
                if (map.TryGetValue(key, out value))
                {
                    return value;
                }

                var normalizedKey = NormalizeLookupKey(key);
                foreach (var pair in map)
                {
                    if (string.Equals(NormalizeLookupKey(pair.Key), normalizedKey, StringComparison.OrdinalIgnoreCase))
                    {
                        return pair.Value;
                    }
                }
            }

            return string.Empty;
        }

        private Dictionary<string, string> LoadBatteryInfoViewMap(string path)
        {
            Dictionary<string, string> bestMap = null;
            var bestScore = -1;

            foreach (var encoding in GetBatteryTextEncodings())
            {
                try
                {
                    var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var line in File.ReadAllLines(path, encoding))
                    {
                        var parts = line.Split(new[] { '\t' }, StringSplitOptions.None);
                        if (parts.Length >= 2)
                        {
                            var key = parts[0].Trim();
                            if (!string.IsNullOrWhiteSpace(key))
                            {
                                map[key] = parts[1].Trim();
                            }
                        }
                    }

                    var score = ScoreBatteryInfoViewMap(map);
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestMap = map;
                    }

                    if (score >= 2)
                    {
                        return map;
                    }
                }
                catch
                {
                }
            }

            return bestMap ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private static int ScoreBatteryInfoViewMap(Dictionary<string, string> map)
        {
            if (map == null || map.Count == 0)
            {
                return 0;
            }

            var score = 0;
            if (ParseBatteryNumber(GetDictionaryValue(map, "电池满电容量", "Full Charged Capacity", "Full Charge Capacity", "完全充電時の容量", "満充電容量")) > 0)
            {
                score++;
            }

            if (ParseBatteryNumber(GetDictionaryValue(map, "电池设计容量", "Designed Capacity", "Design Capacity", "設計容量", "バッテリー設計容量")) > 0)
            {
                score++;
            }

            if (!string.IsNullOrWhiteSpace(GetDictionaryValue(map, "电源状态", "Power State", "電源状態", "電源の状態")))
            {
                score++;
            }

            if (map.Keys.Any(key => !string.IsNullOrWhiteSpace(key) && key.IndexOf('\uFFFD') < 0))
            {
                score++;
            }

            return score;
        }

        private static List<Encoding> GetBatteryTextEncodings()
        {
            var result = new List<Encoding>();
            AddUniqueEncoding(result, new UTF8Encoding(false, true));
            AddUniqueEncoding(result, Encoding.UTF8);
            AddUniqueEncoding(result, Encoding.Unicode);
            AddUniqueEncoding(result, Encoding.BigEndianUnicode);
            AddUniqueEncoding(result, Encoding.Default);

            try
            {
                AddUniqueEncoding(result, Encoding.GetEncoding(932));
            }
            catch
            {
            }

            return result;
        }

        private static void AddUniqueEncoding(List<Encoding> encodings, Encoding encoding)
        {
            if (encodings == null || encoding == null)
            {
                return;
            }

            if (encodings.Any(item => item != null && item.CodePage == encoding.CodePage))
            {
                return;
            }

            encodings.Add(encoding);
        }

        private static Encoding DetectTextFileBomEncoding(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return null;
            }

            try
            {
                using (var stream = File.OpenRead(path))
                {
                    var bom = new byte[4];
                    var read = stream.Read(bom, 0, bom.Length);
                    if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
                    {
                        return new UTF8Encoding(true);
                    }

                    if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
                    {
                        return Encoding.Unicode;
                    }

                    if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
                    {
                        return Encoding.BigEndianUnicode;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static List<Encoding> GetOfficeTextEncodings()
        {
            var result = new List<Encoding>();
            AddUniqueEncoding(result, new UTF8Encoding(false, true));
            AddUniqueEncoding(result, Encoding.UTF8);

            try
            {
                AddUniqueEncoding(result, Encoding.GetEncoding(932));
            }
            catch
            {
            }

            AddUniqueEncoding(result, Encoding.Default);
            AddUniqueEncoding(result, Encoding.Unicode);
            AddUniqueEncoding(result, Encoding.BigEndianUnicode);
            return result;
        }

        private static List<string> ReadAllLinesWithOfficeEncoding(string path, out Encoding detectedEncoding)
        {
            var bomEncoding = DetectTextFileBomEncoding(path);
            if (bomEncoding != null)
            {
                try
                {
                    detectedEncoding = bomEncoding;
                    return File.ReadAllLines(path, bomEncoding).ToList();
                }
                catch
                {
                }
            }

            foreach (var encoding in GetOfficeTextEncodings())
            {
                try
                {
                    detectedEncoding = encoding;
                    return File.ReadAllLines(path, encoding).ToList();
                }
                catch
                {
                }
            }

            detectedEncoding = new UTF8Encoding(false);
            return File.ReadAllLines(path, detectedEncoding).ToList();
        }

        private static string ReadAllTextWithOfficeEncoding(string path, out Encoding detectedEncoding)
        {
            var bomEncoding = DetectTextFileBomEncoding(path);
            if (bomEncoding != null)
            {
                try
                {
                    detectedEncoding = bomEncoding;
                    return File.ReadAllText(path, bomEncoding);
                }
                catch
                {
                }
            }

            foreach (var encoding in GetOfficeTextEncodings())
            {
                try
                {
                    detectedEncoding = encoding;
                    return File.ReadAllText(path, encoding);
                }
                catch
                {
                }
            }

            detectedEncoding = new UTF8Encoding(false);
            return File.ReadAllText(path, detectedEncoding);
        }

        private static int ScaleDialogValue(int value, float scale)
        {
            return (int)Math.Round(value * scale);
        }

        private static Size ScaleDialogSize(Size size, float scale)
        {
            return new Size(
                Math.Max(1, ScaleDialogValue(size.Width, scale)),
                Math.Max(1, ScaleDialogValue(size.Height, scale)));
        }

        private static Padding ScaleDialogPadding(Padding padding, float scale)
        {
            return new Padding(
                ScaleDialogValue(padding.Left, scale),
                ScaleDialogValue(padding.Top, scale),
                ScaleDialogValue(padding.Right, scale),
                ScaleDialogValue(padding.Bottom, scale));
        }

        private static float GetDialogDpiScaleFactor()
        {
            try
            {
                using (var graphics = Graphics.FromHwnd(IntPtr.Zero))
                {
                    return Math.Max(1f, Math.Max(graphics.DpiX, graphics.DpiY) / 96f);
                }
            }
            catch
            {
                return 1f;
            }
        }

        private static void ScaleDialogChildren(Control parent, float scale)
        {
            if (parent == null || scale <= 1.01f)
            {
                return;
            }

            var scrollable = parent as ScrollableControl;
            if (scrollable != null && scrollable.AutoScrollMinSize != Size.Empty)
            {
                scrollable.AutoScrollMinSize = ScaleDialogSize(scrollable.AutoScrollMinSize, scale);
            }

            foreach (Control child in parent.Controls)
            {
                child.Left = ScaleDialogValue(child.Left, scale);
                child.Top = ScaleDialogValue(child.Top, scale);
                child.Width = Math.Max(1, ScaleDialogValue(child.Width, scale));
                child.Height = Math.Max(1, ScaleDialogValue(child.Height, scale));
                child.Margin = ScaleDialogPadding(child.Margin, scale);
                child.Padding = ScaleDialogPadding(child.Padding, scale);

                if (child.Font != null)
                {
                    child.Font = new Font(
                        child.Font.FontFamily,
                        Math.Max(1f, child.Font.Size * scale),
                        child.Font.Style,
                        child.Font.Unit,
                        child.Font.GdiCharSet,
                        child.Font.GdiVerticalFont);
                }

                if (child.Controls.Count > 0)
                {
                    ScaleDialogChildren(child, scale);
                }
            }
        }

        private static void ApplyDialogDpiScaling(Form dialog, Size baseClientSize)
        {
            if (dialog == null)
            {
                return;
            }

            var scale = GetDialogDpiScaleFactor();
            dialog.AutoScaleMode = AutoScaleMode.None;
            dialog.ClientSize = scale <= 1.01f ? baseClientSize : ScaleDialogSize(baseClientSize, scale);
            ScaleDialogChildren(dialog, scale);
        }

        private void ApplyBatteryMetrics(
            double fullChargeCapacity,
            double designCapacity,
            string status,
            Func<double, string> fullChargeFormatter,
            Func<double, string> designFormatter)
        {
            if (fullChargeCapacity <= 0 || designCapacity <= 0)
            {
                SetBatteryUnavailable("未读取到有效电池信息");
                return;
            }

            var health = fullChargeCapacity / designCapacity * 100d;
            var wear = 100d - health;

            _batteryPresent = true;
            _batteryHealthPercent = health;
            _batteryFullChargeCapacityMWh = fullChargeCapacity;
            _batteryRemainingPercent = GetBatteryEstimatedRemainingPercent();
            _batteryHealthLabel.Text = string.Format("{0:0.0}%", health);
            _batteryWearLabel.Text = string.Format("{0:0.0}%", wear < 0 ? 0 : wear);
            _batteryCapacityLabel.Text = string.Format(
                "{0} / {1}",
                fullChargeFormatter == null ? FormatBatteryToolCapacity(fullChargeCapacity) : fullChargeFormatter(fullChargeCapacity),
                designFormatter == null ? FormatBatteryToolCapacity(designCapacity) : designFormatter(designCapacity));
            _batteryStatusLabel.Text = string.IsNullOrWhiteSpace(status) ? "已读取电池信息" : status;
            SyncBatteryActionButtonState();
        }

        private static string NormalizeLookupKey(string value)
        {
            return Regex.Replace(value ?? string.Empty, @"[\s:：_\-／/（）\(\)\[\]【】]", string.Empty).Trim();
        }

        private bool TryDetectBatteryDeviceFast(out bool hasBatteryDevice)
        {
            hasBatteryDevice = true;
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT DeviceID FROM Win32_Battery"))
                {
                    var battery = searcher.Get().Cast<ManagementObject>().FirstOrDefault();
                    hasBatteryDevice = battery != null;
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private void SetBatteryUnavailable(string status)
        {
            _batteryPresent = false;
            _batteryHealthPercent = -1d;
            _batteryFullChargeCapacityMWh = 0d;
            _batteryRemainingPercent = -1d;
            _batteryHealthLabel.Text = "无数据";
            _batteryWearLabel.Text = "无数据";
            _batteryCapacityLabel.Text = "无数据";
            _batteryStatusLabel.Text = status;
            SyncBatteryActionButtonState();
        }

        private static double GetBatteryEstimatedRemainingPercent()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT EstimatedChargeRemaining FROM Win32_Battery"))
                {
                    var battery = searcher.Get().Cast<ManagementObject>().FirstOrDefault();
                    if (battery == null)
                    {
                        return -1d;
                    }

                    var rawValue = SafeToString(battery["EstimatedChargeRemaining"]);
                    double percent;
                    if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out percent)
                        || double.TryParse(rawValue, NumberStyles.Float, CultureInfo.CurrentCulture, out percent))
                    {
                        return percent >= 0d && percent <= 100d ? percent : -1d;
                    }
                }
            }
            catch
            {
            }

            return -1d;
        }

        private void LoadStatusInfo()
        {
            _wifiLabel.Text = BuildWifiSummary();
            _bluetoothLabel.Text = DetectBluetoothStatus();
            _panasonicLabel.Text = DetectPanasonicExtraStatus();
        }

        private void UpdateWifiActionButtonPresentation()
        {
            if (_wifiActionButton == null || _wifiActionButton.IsDisposed)
            {
                return;
            }

            var info = _wifiActionButton.Tag as ButtonStateInfo;
            if (info == null)
            {
                return;
            }

            info.BaseTitle = "1. 联网";
            info.Subtitle = BuildWifiActionSubtitle();
            _toolTip.SetToolTip(_wifiActionButton, BuildWifiActionTooltipText());
            ApplyButtonVisualState(_wifiActionButton);
            UpdateCompactModeWifiButtons();
        }

        private void SyncWifiActionButtonState()
        {
            if (_wifiActionButton == null || _wifiActionButton.IsDisposed)
            {
                return;
            }

            var info = _wifiActionButton.Tag as ButtonStateInfo;
            if (info == null)
            {
                return;
            }

            bool connected;
            string connectedSsid;
            if (TryGetCurrentWifiConnection(out connected, out connectedSsid) && connected)
            {
                SetButtonState(_wifiActionButton, "passed");
            }
            else
            {
                SetButtonState(_wifiActionButton, "failed");
            }

            UpdateWifiActionButtonPresentation();
        }

        private void SyncBatteryActionButtonState()
        {
            if (_batteryActionButton == null || _batteryActionButton.IsDisposed)
            {
                return;
            }

            if (!_batteryPresent || _batteryHealthPercent < 0)
            {
                SetButtonState(_batteryActionButton, "failed");
                return;
            }

            var failByHealth = _settings.BatteryRedHealthPercent > 0 && _batteryHealthPercent < _settings.BatteryRedHealthPercent;
            var failByCapacity = _settings.BatteryRedCapacityMWh > 0 && _batteryFullChargeCapacityMWh < _settings.BatteryRedCapacityMWh;
            if (failByHealth || failByCapacity)
            {
                SetButtonState(_batteryActionButton, "failed");
                return;
            }

            var passByHealth = _settings.BatteryGreenHealthPercent > 0 && _batteryHealthPercent >= _settings.BatteryGreenHealthPercent;
            var passByCapacity = _settings.BatteryGreenCapacityMWh > 0 && _batteryFullChargeCapacityMWh >= _settings.BatteryGreenCapacityMWh;
            var useHealthRule = _settings.BatteryGreenHealthPercent > 0 || _settings.BatteryRedHealthPercent > 0;
            var useCapacityRule = _settings.BatteryGreenCapacityMWh > 0 || _settings.BatteryRedCapacityMWh > 0;

            if ((!useHealthRule || passByHealth) && (!useCapacityRule || passByCapacity))
            {
                SetButtonState(_batteryActionButton, "passed");
                return;
            }

            SetButtonState(_batteryActionButton, "default");
        }

        private bool IsAnyWifiConnected()
        {
            try
            {
                var result = RunCommand("netsh", "wlan show interfaces");
                var lines = result.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string connectedSsid = null;
                var stateConnected = lines.Any(rawLine =>
                {
                    var line = rawLine.Trim();
                    if (!(line.StartsWith("State", StringComparison.OrdinalIgnoreCase)
                        || line.StartsWith("状态", StringComparison.OrdinalIgnoreCase)
                        || line.StartsWith("状態", StringComparison.OrdinalIgnoreCase)))
                    {
                        return false;
                    }

                    return Regex.IsMatch(line, @"^(State|状态|状態)\s*:\s*(connected|已连接|接続済み|接続)\s*$", RegexOptions.IgnoreCase);
                });

                foreach (var rawLine in lines)
                {
                    var line = rawLine.Trim();
                    if (line.IndexOf("BSSID", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        continue;
                    }

                    if (line.StartsWith("SSID", StringComparison.OrdinalIgnoreCase))
                    {
                        var parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            var value = parts[1].Trim();
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                connectedSsid = value;
                                break;
                            }
                        }
                    }
                }

                return stateConnected || !string.IsNullOrWhiteSpace(connectedSsid);
            }
            catch
            {
                return false;
            }
        }

        private string BuildWifiSummary()
        {
            if (_settings.WifiProfiles.Count == 0)
            {
                return "未设置";
            }

            var preview = _settings.WifiProfiles.Take(3).Select((profile, index) =>
            {
                if (string.Equals(profile.Authentication, "open", StringComparison.OrdinalIgnoreCase))
                {
                    return string.Format("{0}.{1}(开放)", index + 1, profile.Ssid);
                }

                return string.Format("{0}.{1}", index + 1, profile.Ssid);
            });

            return string.Join(" -> ", preview.ToArray());
        }

        private string DetectBluetoothStatus()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT Name, Status FROM Win32_PnPEntity WHERE Name LIKE '%Bluetooth%'"))
                {
                    var items = searcher.Get().Cast<ManagementObject>().ToList();
                    if (items.Count == 0)
                    {
                        return "未检测到蓝牙驱动/图标";
                    }

                    var first = items[0];
                    return string.Format("已检测到：{0} | {1}", SafeToString(first["Name"]), SafeToString(first["Status"]));
                }
            }
            catch
            {
                return "检测失败";
            }
        }

        private string DetectPanasonicExtraStatus()
        {
            var isPanasonic = _computerManufacturer.IndexOf("Panasonic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              _computerModel.IndexOf("Panasonic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              _computerModel.IndexOf("CF-", StringComparison.OrdinalIgnoreCase) >= 0;

            if (!isPanasonic)
            {
                return "当前机型非松下，跳过";
            }

            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT Name, Status FROM Win32_PnPEntity WHERE Name LIKE '%Panasonic%' OR Name LIKE '%Wheel%' OR Name LIKE '%Ring%' OR Name LIKE '%Dial%' OR Name LIKE '%Scroll%'"))
                {
                    var items = searcher.Get()
                        .Cast<ManagementObject>()
                        .Where(item =>
                        {
                            var name = SafeToString(item["Name"]);
                            if (string.IsNullOrWhiteSpace(name))
                            {
                                return false;
                            }

                            var upperName = name.ToUpperInvariant();
                            var hasRingKeyword =
                                upperName.Contains("WHEEL") ||
                                upperName.Contains("RING") ||
                                upperName.Contains("DIAL") ||
                                upperName.Contains("SCROLL") ||
                                name.IndexOf("ホイール", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                name.IndexOf("リング", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                name.IndexOf("ダイヤル", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                name.IndexOf("スクロール", StringComparison.OrdinalIgnoreCase) >= 0;
                            var looksLikePlainTouchpad =
                                upperName.Contains("TOUCH PAD") ||
                                upperName.Contains("TOUCHPAD") ||
                                upperName.Contains("MOUSE");

                            return hasRingKeyword && !looksLikePlainTouchpad;
                        })
                        .ToList();
                    if (items.Count == 0)
                    {
                        return "未检测到松下圆环硬件/驱动";
                    }

                    var first = items[0];
                    return string.Format("已检测到：{0} | {1}", SafeToString(first["Name"]), SafeToString(first["Status"]));
                }
            }
            catch
            {
                return "检测失败";
            }
        }

        private static string SafeToString(object value)
        {
            return value == null ? string.Empty : value.ToString().Trim();
        }

        private static string BuildWindowsSummary(string caption, string architecture, string buildNumber)
        {
            var shortCaption = caption.Replace("Microsoft ", string.Empty).Trim();
            var displayVersion = GetWindowsDisplayVersion();

            if (!string.IsNullOrWhiteSpace(displayVersion))
            {
                return string.Format("{0} {1} {2}", shortCaption, displayVersion, architecture).Trim();
            }

            if (!string.IsNullOrWhiteSpace(buildNumber))
            {
                return string.Format("{0} Build {1} {2}", shortCaption, buildNumber, architecture).Trim();
            }

            return string.Format("{0} {1}", shortCaption, architecture).Trim();
        }

        private static string BuildMemorySummary(List<ManagementObject> memories, double totalPhysicalBytes, double visiblePhysicalBytes)
        {
            if (memories == null || memories.Count == 0)
            {
                return FormatTotalWithActualUsable(totalPhysicalBytes, visiblePhysicalBytes);
            }

            var totalFromModules = memories.Sum(memory => ToDouble(memory["Capacity"]));
            var memoryParts = memories.Select((memory, index) =>
            {
                var vendor = SafeToString(memory["Manufacturer"]);
                var capacity = FormatBytes(ToDouble(memory["Capacity"]));
                var slot = SafeToString(memory["DeviceLocator"]);
                if (string.IsNullOrWhiteSpace(slot))
                {
                    slot = SafeToString(memory["BankLabel"]);
                }

                return string.Format(
                    "槽{0} {1} {2}{3}",
                    index + 1,
                    string.IsNullOrWhiteSpace(vendor) ? "未知厂家" : vendor,
                    capacity,
                    string.IsNullOrWhiteSpace(slot) ? string.Empty : string.Format(" ({0})", slot));
            }).ToArray();

            return string.Format(
                "{0}，{1} 条\r\n{2}",
                FormatTotalWithActualUsable(totalFromModules > 0 ? totalFromModules : totalPhysicalBytes, visiblePhysicalBytes),
                memories.Count,
                string.Join("\r\n", memoryParts));
        }

        private static string BuildDiskSummary(List<ManagementObject> disks, List<ManagementObject> logicalDisks)
        {
            if (disks == null || disks.Count == 0)
            {
                return "无数据";
            }

            var totalSize = disks.Sum(d => ToDouble(d["Size"]));
            var totalFree = logicalDisks == null ? 0d : logicalDisks.Sum(d => ToDouble(d["FreeSpace"]));
            var diskParts = disks.Select((disk, index) =>
            {
                var model = SafeToString(disk["Model"]);
                var size = FormatBytes(ToDouble(disk["Size"]));
                var type = DetectDiskType(
                    SafeToString(disk["Model"]),
                    SafeToString(disk["MediaType"]),
                    SafeToString(disk["InterfaceType"]),
                    SafeToString(disk["PNPDeviceID"]));

                return string.Format("{0}. {1} {2} {3}", index + 1, model, size, type).Trim();
            }).ToArray();

            return string.Format("总容量 {0}\r\n{1}", FormatTotalWithActualUsable(totalSize, totalFree), string.Join("\r\n", diskParts));
        }

        private static string FormatTotalWithActualUsable(double totalBytes, double actualUsableBytes)
        {
            if (actualUsableBytes > 0)
            {
                return string.Format("{0}（实际可用 {1}）", FormatBytes(totalBytes), FormatBytes(actualUsableBytes));
            }

            return FormatBytes(totalBytes);
        }

        private static double GetVisiblePhysicalMemoryBytes()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT TotalVisibleMemorySize FROM Win32_OperatingSystem"))
                {
                    var os = searcher.Get().Cast<ManagementObject>().FirstOrDefault();
                    if (os == null)
                    {
                        return 0d;
                    }

                    return ToDouble(os["TotalVisibleMemorySize"]) * 1024d;
                }
            }
            catch
            {
                return 0d;
            }
        }

        private static string DetectDiskType(string model, string mediaType, string interfaceType, string pnpDeviceId)
        {
            var upperModel = (model ?? string.Empty).ToUpperInvariant();
            var upperMedia = (mediaType ?? string.Empty).ToUpperInvariant();
            var upperInterface = (interfaceType ?? string.Empty).ToUpperInvariant();
            var upperPnp = (pnpDeviceId ?? string.Empty).ToUpperInvariant();

            if (upperModel.Contains("NVME") || upperPnp.Contains("NVME"))
            {
                return "NVMe / M.2";
            }

            if (upperMedia.Contains("SSD") || upperModel.Contains("SSD"))
            {
                return upperInterface.Contains("USB") ? "USB SSD" : "SATA SSD";
            }

            if (upperMedia.Contains("HARD") || upperMedia.Contains("FIXED") || upperModel.Contains("HDD"))
            {
                return upperInterface.Contains("USB") ? "USB 机械盘" : "机械盘";
            }

            if (upperInterface.Contains("USB"))
            {
                return "USB 磁盘";
            }

            if (upperInterface.Contains("IDE") || upperInterface.Contains("SATA"))
            {
                return "SATA 磁盘";
            }

            return string.IsNullOrWhiteSpace(interfaceType) ? "磁盘类型待识别" : interfaceType;
        }

        private static string GetWindowsDisplayVersion()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    if (key == null)
                    {
                        return string.Empty;
                    }

                    var displayVersion = SafeToString(key.GetValue("DisplayVersion"));
                    if (!string.IsNullOrWhiteSpace(displayVersion))
                    {
                        return displayVersion;
                    }

                    return SafeToString(key.GetValue("ReleaseId"));
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static double ToDouble(object value)
        {
            double result;
            return double.TryParse(SafeToString(value), out result) ? result : 0d;
        }

        private static string FormatBytes(double bytes)
        {
            if (bytes <= 0)
            {
                return "无数据";
            }

            var gb = bytes / 1024d / 1024d / 1024d;
            if (gb >= 1024d)
            {
                return string.Format("{0:0.00} TB", gb / 1024d);
            }

            return string.Format("{0:0.##} GB", gb);
        }

        private static string FormatMilliwattHour(double value)
        {
            if (value <= 0)
            {
                return "无数据";
            }

            return string.Format("{0:0.##} Wh", value / 1000d);
        }

        private static string FormatBatteryToolCapacity(double value)
        {
            if (value <= 0)
            {
                return "无数据";
            }

            return string.Format("{0:0} mWh", value);
        }

        private static double ParseBatteryNumber(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0d;
            }

            var match = Regex.Match(text.Replace(",", string.Empty), @"([0-9]+(?:\.[0-9]+)?)");
            if (!match.Success)
            {
                return 0d;
            }

            double value;
            return double.TryParse(match.Groups[1].Value, out value) ? value : 0d;
        }

        private static double ParseBatteryPercent(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return -1d;
            }

            var match = Regex.Match(text, @"([0-9]+(?:\.[0-9]+)?)");
            if (!match.Success)
            {
                return -1d;
            }

            double value;
            return double.TryParse(match.Groups[1].Value, out value) ? value : -1d;
        }

        private static double ExtractBatteryValue(string html, params string[] labels)
        {
            var labelPattern = BuildBatteryLabelRegex(labels);
            if (string.IsNullOrWhiteSpace(labelPattern))
            {
                return 0d;
            }

            var match = Regex.Match(
                html,
                string.Format("(?:{0})</span></td><td><span class=\"value\">([0-9,]+) mWh", labelPattern),
                RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                return 0d;
            }

            double value;
            return double.TryParse(match.Groups[1].Value.Replace(",", string.Empty), out value) ? value : 0d;
        }

        private static string ExtractBatteryText(string html, params string[] labels)
        {
            var labelPattern = BuildBatteryLabelRegex(labels);
            if (string.IsNullOrWhiteSpace(labelPattern))
            {
                return string.Empty;
            }

            var match = Regex.Match(
                html,
                string.Format("(?:{0})</span></td><td><span class=\"value\">([^<]+)", labelPattern),
                RegexOptions.IgnoreCase);

            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }

        private static string BuildBatteryLabelRegex(IEnumerable<string> labels)
        {
            if (labels == null)
            {
                return string.Empty;
            }

            var patterns = labels
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .Select(label => Regex.Escape(label).Replace("\\ ", "\\s*"))
                .Distinct()
                .ToArray();
            return patterns.Length == 0 ? string.Empty : string.Join("|", patterns);
        }

        private void ShowInfo(string message, string title = null)
        {
            ShowCopyableMessage(message, string.IsNullOrWhiteSpace(title) ? GetConfiguredAppTitle() : title, false);
        }

        private void ShowErrorMessage(string message, string title = null)
        {
            ShowCopyableMessage(message, string.IsNullOrWhiteSpace(title) ? GetConfiguredAppTitle() : title, true);
        }

        private void ShowCopyableMessage(string message, string title, bool isError)
        {
            using (var dialog = new Form())
            {
                dialog.Text = title;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.ClientSize = new Size(560, 280);
                dialog.Font = Font;
                dialog.BackColor = isError ? Color.FromArgb(254, 242, 242) : Color.FromArgb(240, 249, 255);

                var titleBox = new TextBox
                {
                    Text = isError ? "提示信息（可复制）" : "执行结果（可复制）",
                    Location = new Point(20, 18),
                    Size = new Size(240, 28),
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    BackColor = dialog.BackColor,
                    ForeColor = isError ? Color.FromArgb(153, 27, 27) : Color.FromArgb(3, 105, 161),
                    Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold, GraphicsUnit.Point),
                    TabStop = false
                };

                var messageBox = new TextBox
                {
                    Text = message,
                    Location = new Point(20, 56),
                    Size = new Size(520, 168),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    ShortcutsEnabled = true,
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = Color.White,
                    ForeColor = Color.FromArgb(15, 23, 42)
                };

                var copyButton = new Button
                {
                    Text = "复制",
                    Location = new Point(356, 236),
                    Size = new Size(84, 30)
                };

                var okButton = new Button
                {
                    Text = "确定",
                    Location = new Point(456, 236),
                    Size = new Size(84, 30)
                };

                copyButton.Click += (_, __) =>
                {
                    try
                    {
                        Clipboard.SetText(messageBox.Text ?? string.Empty);
                    }
                    catch
                    {
                    }
                };

                okButton.Click += (_, __) => dialog.Close();
                dialog.AcceptButton = okButton;
                dialog.Controls.Add(titleBox);
                dialog.Controls.Add(messageBox);
                dialog.Controls.Add(copyButton);
                dialog.Controls.Add(okButton);
                dialog.ShowDialog(this);
            }
        }

        private bool TryStartTarget(string pathOrUri, string arguments = null)
        {
            if (string.IsNullOrWhiteSpace(pathOrUri))
            {
                return false;
            }

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = pathOrUri,
                    UseShellExecute = true
                };

                if (!string.IsNullOrWhiteSpace(arguments))
                {
                    startInfo.Arguments = arguments;
                }

                if (!IsShellUri(pathOrUri) && Path.IsPathRooted(pathOrUri) && File.Exists(pathOrUri))
                {
                    var workingDirectory = Path.GetDirectoryName(pathOrUri);
                    if (!string.IsNullOrWhiteSpace(workingDirectory))
                    {
                        startInfo.WorkingDirectory = workingDirectory;
                    }
                }

                Process.Start(startInfo);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool StartTarget(string pathOrUri, string arguments = null)
        {
            if (TryStartTarget(pathOrUri, arguments))
            {
                return true;
            }

            ShowErrorMessage(string.Format("打开失败：{0}", pathOrUri));
            return false;
        }

        private string RunCommand(string fileName, string arguments)
        {
            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                process.Start();
                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(error) ? output : error);
                }

                return string.IsNullOrWhiteSpace(output) ? error : output;
            }
        }

        private void SaveWifiProfile(string ssid, string password, string authentication, string encryption)
        {
            var safeSsid = SecurityElement.Escape(ssid);
            var isOpen = string.Equals(authentication, "open", StringComparison.OrdinalIgnoreCase);
            string profileXml;

            if (isOpen)
            {
                profileXml =
                    "<?xml version=\"1.0\"?>" +
                    "<WLANProfile xmlns=\"http://www.microsoft.com/networking/WLAN/profile/v1\">" +
                    string.Format("<name>{0}</name>", safeSsid) +
                    "<SSIDConfig><SSID>" +
                    string.Format("<name>{0}</name>", safeSsid) +
                    "</SSID></SSIDConfig>" +
                    "<connectionType>ESS</connectionType>" +
                    "<connectionMode>auto</connectionMode>" +
                    "<MSM><security><authEncryption>" +
                    "<authentication>open</authentication>" +
                    "<encryption>none</encryption>" +
                    "<useOneX>false</useOneX>" +
                    "</authEncryption></security></MSM>" +
                    "</WLANProfile>";
            }
            else
            {
                var safePassword = SecurityElement.Escape(password);
                profileXml =
                    "<?xml version=\"1.0\"?>" +
                    "<WLANProfile xmlns=\"http://www.microsoft.com/networking/WLAN/profile/v1\">" +
                    string.Format("<name>{0}</name>", safeSsid) +
                    "<SSIDConfig><SSID>" +
                    string.Format("<name>{0}</name>", safeSsid) +
                    "</SSID></SSIDConfig>" +
                    "<connectionType>ESS</connectionType>" +
                    "<connectionMode>auto</connectionMode>" +
                    "<MSM><security><authEncryption>" +
                    string.Format("<authentication>{0}</authentication>", authentication) +
                    string.Format("<encryption>{0}</encryption>", encryption) +
                    "<useOneX>false</useOneX>" +
                    "</authEncryption><sharedKey>" +
                    "<keyType>passPhrase</keyType>" +
                    "<protected>false</protected>" +
                    string.Format("<keyMaterial>{0}</keyMaterial>", safePassword) +
                    "</sharedKey></security></MSM>" +
                    "</WLANProfile>";
            }

            var tempFile = Path.Combine(Path.GetTempPath(), string.Format("wifi-profile-{0}.xml", Guid.NewGuid().ToString("N")));
            File.WriteAllText(tempFile, profileXml, new UTF8Encoding(false));

            try
            {
                RunCommand("netsh", string.Format("wlan add profile filename=\"{0}\" user=all", tempFile));
            }
            finally
            {
                try
                {
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }
                catch
                {
                }
            }
        }

        private void ConnectWifi(string ssid, string password, string authentication, string encryption)
        {
            if (IsConnectedToSsid(ssid))
            {
                return;
            }

            EnsureWirelessInterfaceEnabled();
            WaitForWirelessReady(9000);
            SaveWifiProfile(ssid, password, authentication, encryption);
            WaitForWifiProfileRegistration(ssid, 5000);

            if (IsConnectedToSsid(ssid))
            {
                return;
            }

            var previousConnectedSsid = GetCurrentConnectedWifiSsid();
            DisconnectWirelessInterfaces();
            System.Threading.Thread.Sleep(900);

            Exception lastError = null;
            string lastCommandError = null;
            for (var i = 0; i < 4; i++)
            {
                try
                {
                    EnsureWirelessInterfaceEnabled();
                    WaitForWirelessReady(7000);
                    lastCommandError = TryConnectWifiUsingKnownInterfaces(ssid);

                    if (WaitForWifiConnection(ssid, previousConnectedSsid, 16000))
                    {
                        return;
                    }

                    if (i < 3)
                    {
                        DisconnectWirelessInterfaces();
                        System.Threading.Thread.Sleep(1200 + (i * 300));
                    }
                }
                catch (Exception ex)
                {
                    lastError = ex;
                }
            }

            if (WaitForWifiConnection(ssid, previousConnectedSsid, 8000))
            {
                return;
            }

            if (lastError != null)
            {
                throw lastError;
            }

            if (!string.IsNullOrWhiteSpace(lastCommandError))
            {
                throw new InvalidOperationException(lastCommandError);
            }

            throw new InvalidOperationException(string.Format("未成功连接到 WiFi：{0}", ssid));
        }

        private string TryConnectWifiUsingKnownInterfaces(string ssid)
        {
            var interfaceNames = GetWirelessInterfaceNames();
            string lastError = null;

            if (interfaceNames.Count == 0)
            {
                return TryConnectWifiWithNetsh(ssid, null);
            }

            foreach (var interfaceName in interfaceNames)
            {
                TryRunCommand("netsh", string.Format("wlan set autoconfig enabled=yes interface=\"{0}\"", interfaceName));
                var interfaceError = TryConnectWifiWithNetsh(ssid, interfaceName);
                if (string.IsNullOrWhiteSpace(interfaceError))
                {
                    return string.Empty;
                }

                lastError = interfaceError;
            }

            var genericError = TryConnectWifiWithNetsh(ssid, null);
            return string.IsNullOrWhiteSpace(genericError) ? string.Empty : (lastError ?? genericError);
        }

        private string TryConnectWifiWithNetsh(string ssid, string interfaceName)
        {
            var arguments = string.IsNullOrWhiteSpace(interfaceName)
                ? string.Format("wlan connect name=\"{0}\" ssid=\"{0}\"", ssid)
                : string.Format("wlan connect name=\"{0}\" ssid=\"{0}\" interface=\"{1}\"", ssid, interfaceName);
            var result = RunProcessForOutputDetailed("netsh", arguments, GetApplicationDirectory(), 10000);
            if (!result.Started)
            {
                return "联网命令未能启动。";
            }

            if (result.TimedOut)
            {
                return "联网命令执行超时。";
            }

            if (result.ExitCode == 0)
            {
                return string.Empty;
            }

            if (WasWifiConnectRequestAccepted(result))
            {
                return string.Empty;
            }

            var firstLine = ExtractFirstMeaningfulLine(result.CombinedOutput);
            return string.IsNullOrWhiteSpace(firstLine) ? "联网命令执行失败。" : firstLine;
        }

        private bool WasWifiConnectRequestAccepted(ProcessExecutionResult result)
        {
            if (result == null || string.IsNullOrWhiteSpace(result.CombinedOutput))
            {
                return false;
            }

            return ContainsAny(
                result.CombinedOutput,
                "Connection request was completed successfully",
                "连接请求已成功完成",
                "已成功完成连接请求",
                "接続要求は正常に完了",
                "接続要求が正常に完了",
                "接続要求は正常に完了しました",
                "接続要求が正常に完了しました");
        }

        private void WaitForWifiProfileRegistration(string ssid, int timeoutMs)
        {
            if (string.IsNullOrWhiteSpace(ssid))
            {
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                if (IsWifiProfileRegistered(ssid))
                {
                    return;
                }

                System.Threading.Thread.Sleep(350);
            }
        }

        private bool IsWifiProfileRegistered(string ssid)
        {
            try
            {
                var output = RunCommand("netsh", "wlan show profiles");
                return output.IndexOf(ssid, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private bool WaitForWirelessReady(int timeoutMs)
        {
            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                var status = GetWirelessEnvironmentStatus(false);
                if (status.CanAttemptConnect)
                {
                    return true;
                }

                System.Threading.Thread.Sleep(400);
            }

            return false;
        }

        private void DisconnectWirelessInterfaces()
        {
            foreach (var interfaceName in GetWirelessInterfaceNames())
            {
                TryRunCommand("netsh", string.Format("wlan disconnect interface=\"{0}\"", interfaceName));
            }
        }

        private List<string> GetWirelessInterfaceNames()
        {
            var result = new List<string>();

            try
            {
                var output = RunCommand("netsh", "wlan show interfaces");
                var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var rawLine in lines)
                {
                    var line = rawLine.Trim();
                    if (line.StartsWith("Name", StringComparison.OrdinalIgnoreCase)
                        || line.StartsWith("名称", StringComparison.OrdinalIgnoreCase)
                        || line.StartsWith("名前", StringComparison.OrdinalIgnoreCase))
                    {
                        var parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            var name = parts[1].Trim();
                            if (!string.IsNullOrWhiteSpace(name) && !result.Contains(name))
                            {
                                result.Add(name);
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            try
            {
                var output = RunCommand("netsh", "interface show interface");
                var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var rawLine in lines.Skip(2))
                {
                    var line = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    if (line.Contains("WLAN") || line.Contains("Wi-Fi") || line.Contains("无线") || line.Contains("ワイヤレス"))
                    {
                        var parts = Regex.Split(line, "\\s{2,}");
                        if (parts.Length > 0)
                        {
                            var name = parts[parts.Length - 1].Trim();
                            if (!string.IsNullOrWhiteSpace(name) && !result.Contains(name))
                            {
                                result.Add(name);
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            return result;
        }

        private void EnsureWirelessInterfaceEnabled()
        {
            try
            {
                TryRunCommand("sc", "start WlanSvc");

                using (var searcher = new ManagementObjectSearcher(
                    "SELECT Name, NetConnectionID, NetEnabled, Index FROM Win32_NetworkAdapter WHERE PhysicalAdapter = True"))
                {
                    var adapters = searcher.Get().Cast<ManagementObject>().Where(adapter =>
                    {
                        var name = SafeToString(adapter["Name"]);
                        var connectionId = SafeToString(adapter["NetConnectionID"]);
                        var combined = (name + " " + connectionId).ToUpperInvariant();
                        return combined.Contains("WI-FI") || combined.Contains("WIFI") || combined.Contains("WIRELESS") || combined.Contains("802.11") || combined.Contains("无线") || combined.Contains("ワイヤレス");
                    }).ToList();

                    foreach (var adapter in adapters)
                    {
                        var connectionId = SafeToString(adapter["NetConnectionID"]);
                        try
                        {
                            adapter.InvokeMethod("Enable", null);
                        }
                        catch
                        {
                        }

                        if (!string.IsNullOrWhiteSpace(connectionId))
                        {
                            TryRunCommand("netsh", string.Format("interface set interface name=\"{0}\" admin=enabled", connectionId));
                        }
                    }
                }

                TryTurnOnWirelessRadio();
                System.Threading.Thread.Sleep(1800);
            }
            catch
            {
            }
        }

        private bool WaitForWirelessAdapterActivation(int timeoutMs)
        {
            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    var adapters = GetWirelessAdapters();
                    if (adapters.Any(IsAdapterEnabled))
                    {
                        return true;
                    }
                }
                catch
                {
                }

                try
                {
                    if (GetWirelessInterfaceNames().Count > 0)
                    {
                        return true;
                    }
                }
                catch
                {
                }

                System.Threading.Thread.Sleep(500);
            }

            return false;
        }

        private void TryTurnOnWirelessRadio()
        {
            IntPtr clientHandle = IntPtr.Zero;
            try
            {
                uint negotiatedVersion;
                var openResult = WlanOpenHandle(2, IntPtr.Zero, out negotiatedVersion, out clientHandle);
                if (openResult != 0 || clientHandle == IntPtr.Zero)
                {
                    return;
                }

                foreach (var interfaceGuid in GetWirelessInterfaceGuids())
                {
                    TryTurnOnWirelessRadio(clientHandle, interfaceGuid);
                }
            }
            catch
            {
            }
            finally
            {
                if (clientHandle != IntPtr.Zero)
                {
                    try
                    {
                        WlanCloseHandle(clientHandle, IntPtr.Zero);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private void TryTurnOnWirelessRadio(IntPtr clientHandle, Guid interfaceGuid)
        {
            IntPtr dataPointer = IntPtr.Zero;
            try
            {
                int dataSize;
                var queryResult = WlanQueryInterface(
                    clientHandle,
                    ref interfaceGuid,
                    WlanIntfOpcode.RadioState,
                    IntPtr.Zero,
                    out dataSize,
                    out dataPointer,
                    IntPtr.Zero);

                if (queryResult != 0 || dataPointer == IntPtr.Zero)
                {
                    return;
                }

                var radioState = (WlanRadioState)Marshal.PtrToStructure(dataPointer, typeof(WlanRadioState));
                if (radioState.PhyRadioState == null)
                {
                    return;
                }

                var count = Math.Min((int)radioState.NumberOfPhys, radioState.PhyRadioState.Length);
                for (var i = 0; i < count; i++)
                {
                    var phyState = radioState.PhyRadioState[i];
                    if (phyState.SoftwareRadioState == Dot11RadioState.On)
                    {
                        continue;
                    }

                    phyState.SoftwareRadioState = Dot11RadioState.On;
                    WlanSetInterface(
                        clientHandle,
                        ref interfaceGuid,
                        WlanIntfOpcode.RadioState,
                        Marshal.SizeOf(typeof(WlanPhyRadioState)),
                        ref phyState,
                        IntPtr.Zero);
                }
            }
            catch
            {
            }
            finally
            {
                if (dataPointer != IntPtr.Zero)
                {
                    try
                    {
                        WlanFreeMemory(dataPointer);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private List<Guid> GetWirelessInterfaceGuids()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT GUID, Name, NetConnectionID, NetEnabled, PhysicalAdapter, AdapterType, PNPDeviceID FROM Win32_NetworkAdapter WHERE PhysicalAdapter = True"))
                {
                    return searcher.Get()
                        .Cast<ManagementObject>()
                        .Where(IsWirelessAdapter)
                        .Select(adapter => SafeToString(adapter["GUID"]))
                        .Where(guidText => !string.IsNullOrWhiteSpace(guidText))
                        .Select(guidText =>
                        {
                            Guid guid;
                            return Guid.TryParse(guidText.Trim('{', '}'), out guid) ? guid : Guid.Empty;
                        })
                        .Where(guid => guid != Guid.Empty)
                        .Distinct()
                        .ToList();
                }
            }
            catch
            {
                return new List<Guid>();
            }
        }

        private bool IsConnectedToSsid(string ssid)
        {
            bool connected;
            string connectedSsid;
            string connectedProfile;
            if (!TryGetCurrentWifiConnectionDetails(out connected, out connectedSsid, out connectedProfile) || !connected)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(ssid))
            {
                return connected;
            }

            return MatchesWifiIdentity(ssid, connectedSsid, connectedProfile);
        }

        private string GetCurrentConnectedWifiSsid()
        {
            bool connected;
            string connectedSsid;
            string connectedProfile;
            return TryGetCurrentWifiConnectionDetails(out connected, out connectedSsid, out connectedProfile) && connected
                ? (string.IsNullOrWhiteSpace(connectedSsid) ? connectedProfile : connectedSsid)
                : string.Empty;
        }

        private bool TryGetCurrentWifiConnection(out bool connected, out string connectedSsid)
        {
            string connectedProfile;
            return TryGetCurrentWifiConnectionDetails(out connected, out connectedSsid, out connectedProfile);
        }

        private bool TryGetCurrentWifiConnectionDetails(out bool connected, out string connectedSsid, out string connectedProfile)
        {
            connected = false;
            connectedSsid = string.Empty;
            connectedProfile = string.Empty;
            try
            {
                var result = RunCommand("netsh", "wlan show interfaces");
                var lines = result.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var rawLine in lines)
                {
                    var line = rawLine.Trim();
                    if (line.IndexOf("BSSID", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        continue;
                    }

                    if (Regex.IsMatch(line, @"^(State|状态|状態)\s*:\s*(connected|已连接|接続済み|接続)\s*$", RegexOptions.IgnoreCase))
                    {
                        connected = true;
                        continue;
                    }

                    if (Regex.IsMatch(line, @"^SSID(\s+\d+)?\s*:", RegexOptions.IgnoreCase))
                    {
                        var parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            var value = parts[1].Trim();
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                connectedSsid = value;
                            }
                        }
                    }

                    if (Regex.IsMatch(line, @"^(Profile|配置文件|プロファイル)\s*:", RegexOptions.IgnoreCase))
                    {
                        var parts = line.Split(new[] { ':' }, 2);
                        if (parts.Length == 2)
                        {
                            var value = parts[1].Trim();
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                connectedProfile = value;
                            }
                        }
                    }
                }

                connected = connected || !string.IsNullOrWhiteSpace(connectedSsid) || !string.IsNullOrWhiteSpace(connectedProfile);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool WaitForWifiConnection(string expectedSsid, string previousConnectedSsid, int timeoutMs)
        {
            var stopwatch = Stopwatch.StartNew();

            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                bool connected;
                string connectedSsid;
                string connectedProfile;
                if (TryGetCurrentWifiConnectionDetails(out connected, out connectedSsid, out connectedProfile) && connected)
                {
                    if (MatchesWifiIdentity(expectedSsid, connectedSsid, connectedProfile))
                    {
                        return true;
                    }

                    if (DoesCurrentWifiOutputMentionSsid(expectedSsid))
                    {
                        return true;
                    }
                }

                System.Threading.Thread.Sleep(650);
            }

            return false;
        }

        private bool WaitForAnySavedWifiConnection(IEnumerable<WifiProfile> profiles, int timeoutMs, out WifiProfile matchedProfile)
        {
            matchedProfile = null;
            var candidates = (profiles ?? Enumerable.Empty<WifiProfile>())
                .Where(profile => profile != null && !string.IsNullOrWhiteSpace(profile.Ssid))
                .ToList();
            if (candidates.Count == 0)
            {
                return false;
            }

            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                matchedProfile = candidates.FirstOrDefault(profile => IsConnectedToSsid(profile.Ssid));
                if (matchedProfile != null)
                {
                    return true;
                }

                System.Threading.Thread.Sleep(500);
            }

            return false;
        }

        private bool MatchesWifiIdentity(string expectedSsid, string connectedSsid, string connectedProfile)
        {
            if (string.IsNullOrWhiteSpace(expectedSsid))
            {
                return false;
            }

            return string.Equals(connectedSsid, expectedSsid, StringComparison.OrdinalIgnoreCase)
                || string.Equals(connectedProfile, expectedSsid, StringComparison.OrdinalIgnoreCase);
        }

        private bool DoesCurrentWifiOutputMentionSsid(string expectedSsid)
        {
            if (string.IsNullOrWhiteSpace(expectedSsid))
            {
                return false;
            }

            try
            {
                var output = RunCommand("netsh", "wlan show interfaces");
                return output.IndexOf(expectedSsid, StringComparison.OrdinalIgnoreCase) >= 0
                    && Regex.IsMatch(output, "State\\s*:\\s*connected|状态\\s*:\\s*已连接|状態\\s*:\\s*接続", RegexOptions.IgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private void TryRunCommand(string fileName, string arguments)
        {
            try
            {
                RunCommand(fileName, arguments);
            }
            catch
            {
            }
        }

        private void ConnectWifiFromSettings(Button sourceButton)
        {
            ConnectWifiProfiles(_settings.WifiProfiles, sourceButton, false);
        }

        private void ConnectSpecificWifiProfile(WifiProfile profile, Button sourceButton)
        {
            if (profile == null)
            {
                return;
            }

            ConnectWifiProfiles(new[] { profile }, sourceButton, true);
        }

        private void ConnectWifiProfiles(IEnumerable<WifiProfile> profiles, Button sourceButton, bool singleChoiceMode)
        {
            if (!IsRunningAsAdministrator())
            {
                var result = MessageBox.Show(
                    this,
                    "当前系统的 WiFi 自动连接需要管理员权限。\r\n是否现在以管理员身份重新打开本工具？",
                    "联网需要管理员权限",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    RestartAsAdministrator();
                }

                return;
            }

            var orderedProfiles = (profiles ?? Enumerable.Empty<WifiProfile>())
                .Where(profile => profile != null && !string.IsNullOrWhiteSpace(profile.Ssid))
                .ToList();
            if (orderedProfiles.Count == 0)
            {
                ShowErrorMessage("请先在设置里添加 WiFi 列表。", "联网");
                return;
            }

            var wirelessStatus = GetWirelessEnvironmentStatus(true);
            if (!wirelessStatus.CanAttemptConnect)
            {
                ShowErrorMessage(wirelessStatus.Message, "联网");
                return;
            }

            var errors = new List<string>();
            foreach (var profile in orderedProfiles)
            {
                try
                {
                    ConnectWifi(profile.Ssid, profile.Password, profile.Authentication, profile.Encryption);
                    RefreshDashboard();
                    if (sourceButton != null && !sourceButton.IsDisposed)
                    {
                        SetButtonState(sourceButton, "passed");
                    }
                    return;
                }
                catch (Exception ex)
                {
                    errors.Add(string.Format("{0}: {1}", profile.Ssid, ex.Message));
                    if (IsConnectedToSsid(profile.Ssid))
                    {
                        RefreshDashboard();
                        if (sourceButton != null && !sourceButton.IsDisposed)
                        {
                            SetButtonState(sourceButton, "passed");
                        }
                        return;
                    }
                }
            }

            WifiProfile lateConnectedProfile;
            if (WaitForAnySavedWifiConnection(orderedProfiles, 7000, out lateConnectedProfile))
            {
                RefreshDashboard();
                if (sourceButton != null && !sourceButton.IsDisposed)
                {
                    SetButtonState(sourceButton, "passed");
                }
                return;
            }

            var statusAfterFailure = GetWirelessEnvironmentStatus(false);
            var builder = new StringBuilder();
            builder.AppendLine(singleChoiceMode ? "所选 WiFi 连接失败。" : "设置里的 WiFi 都连接失败。");
            builder.AppendLine();
            builder.AppendLine(singleChoiceMode ? "已按你选择的 WiFi 进行连接。" : "已按设置顺序自动联网。");
            builder.AppendLine("如果名称和密码确认无误，请检查热点是否可见、距离是否过远。");

            if (!statusAfterFailure.CanAttemptConnect)
            {
                builder.AppendLine();
                builder.AppendLine(statusAfterFailure.Message);
            }
            else if (errors.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("失败详情：");
                foreach (var error in errors.Take(3))
                {
                    builder.AppendLine(error);
                }
            }

            ShowErrorMessage(builder.ToString().Trim(), "联网");
        }

        private WirelessEnvironmentStatus GetWirelessEnvironmentStatus(bool tryEnable)
        {
            var status = new WirelessEnvironmentStatus();
            var adapters = GetWirelessAdapters();
            status.HasWirelessAdapter = adapters.Count > 0;

            if (!status.HasWirelessAdapter)
            {
                status.Message = "未检测到无线网卡，无法自动联网。\r\n请先安装无线网卡驱动，或确认这台电脑本身带有 WLAN 网卡。";
                return status;
            }

            status.HasEnabledAdapter = adapters.Any(IsAdapterEnabled);

            if (tryEnable && !status.HasEnabledAdapter)
            {
                EnsureWirelessInterfaceEnabled();
                WaitForWirelessAdapterActivation(9000);
                adapters = GetWirelessAdapters();
                status.HasEnabledAdapter = adapters.Any(IsAdapterEnabled);
            }

            status.InterfaceNames = GetWirelessInterfaceNames();
            status.HasInterface = status.InterfaceNames.Count > 0;
            if (!status.HasEnabledAdapter && status.HasInterface)
            {
                status.HasEnabledAdapter = true;
            }

            string wlanOutput = string.Empty;
            try
            {
                wlanOutput = RunCommand("netsh", "wlan show interfaces");
            }
            catch (Exception ex)
            {
                status.CommandError = ex.Message;
            }

            if (!string.IsNullOrWhiteSpace(wlanOutput))
            {
                status.SoftwareRadioOff = Regex.IsMatch(wlanOutput, "Software\\s+Off|软件.*关闭|软件.*关|ソフトウェア.*オフ|ソフトウェア.*切", RegexOptions.IgnoreCase);
                status.HardwareRadioOff = Regex.IsMatch(wlanOutput, "Hardware\\s+Off|硬件.*关闭|硬件.*关|ハードウェア.*オフ|ハードウェア.*切", RegexOptions.IgnoreCase);

                if (ContainsNoWirelessInterfaceMessage(wlanOutput))
                {
                    status.HasInterface = false;
                    status.InterfaceNames.Clear();
                }
            }

            if (!status.HasEnabledAdapter)
            {
                status.Message = "检测到无线网卡，但当前处于禁用状态，软件已尝试自动启用。\r\n如果仍然不行，请检查设备管理器，或确认机身无线开关没有关闭。";
                return status;
            }

            if (status.HardwareRadioOff)
            {
                status.Message = "检测到无线网卡，但无线硬件开关处于关闭状态。\r\n软件已尝试自动开启 WLAN，仍未成功，请打开机身无线开关后，再点一次“1. 联网”。";
                return status;
            }

            if (status.SoftwareRadioOff)
            {
                status.Message = "检测到无线网卡，但系统 WLAN 开关仍处于关闭状态。\r\n软件已尝试自动打开它；如果还是不行，通常是飞行模式开启，或系统限制了无线开关，请关闭后重试。";
                return status;
            }

            if (!status.HasInterface)
            {
                status.Message = "检测到无线网卡，但系统当前没有可用的 WLAN 接口。\r\n可能是飞行模式、无线服务异常，或网卡驱动异常，请检查后重试。";
                return status;
            }

            status.CanAttemptConnect = true;
            return status;
        }

        private List<ManagementObject> GetWirelessAdapters()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT Name, NetConnectionID, NetEnabled, PhysicalAdapter, AdapterType, PNPDeviceID FROM Win32_NetworkAdapter WHERE PhysicalAdapter = True"))
                {
                    return searcher.Get().Cast<ManagementObject>().Where(IsWirelessAdapter).ToList();
                }
            }
            catch
            {
                return new List<ManagementObject>();
            }
        }

        private static bool IsWirelessAdapter(ManagementObject adapter)
        {
            if (adapter == null)
            {
                return false;
            }

            var name = SafeToString(adapter["Name"]);
            var connectionId = SafeToString(adapter["NetConnectionID"]);
            var adapterType = SafeToString(adapter["AdapterType"]);
            var pnpDeviceId = SafeToString(adapter["PNPDeviceID"]);
            var combined = (name + " " + connectionId + " " + adapterType + " " + pnpDeviceId).ToUpperInvariant();

            return combined.Contains("WI-FI")
                || combined.Contains("WIFI")
                || combined.Contains("WIRELESS")
                || combined.Contains("802.11")
                || combined.Contains("WLAN")
                || combined.Contains("无线")
                || combined.Contains("ワイヤレス");
        }

        private static bool IsAdapterEnabled(ManagementObject adapter)
        {
            if (adapter == null)
            {
                return false;
            }

            var enabledObject = adapter["NetEnabled"];
            if (enabledObject is bool)
            {
                return (bool)enabledObject;
            }

            return string.Equals(SafeToString(enabledObject), "True", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsNoWirelessInterfaceMessage(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return text.IndexOf("There is no wireless interface on the system", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("系统上没有无线接口", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("没有无线接口", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("ワイヤレス インターフェイスがありません", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("システムにワイヤレス インターフェイスはありません", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool IsRunningAsAdministrator()
        {
            try
            {
                var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private void RestartAsAdministrator()
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = Application.ExecutablePath,
                    UseShellExecute = true,
                    Verb = "runas"
                };

                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                ShowErrorMessage(string.Format("无法以管理员身份重新打开：{0}", ex.Message), "联网");
            }
        }

        private void ShowSettingsDialog()
        {
            using (var dialog = new Form())
            {
                dialog.Text = "设置";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.AutoScaleMode = AutoScaleMode.Dpi;
                dialog.AutoScroll = false;
                dialog.ClientSize = new Size(820, 820);
                dialog.Font = Font;
                var tabs = new TabControl
                {
                    Location = new Point(18, 18),
                    Size = new Size(784, 720),
                    Multiline = false
                };

                var wifiTab = new TabPage("WiFi");
                var batteryTab = new TabPage("电池规则");
                var toolTab = new TabPage("工具路径");
                var officePostTab = new TabPage("Office按键");
                var officeCommTab = new TabPage("Office通信");
                var aboutTab = new TabPage("关于");
                wifiTab.AutoScroll = true;
                batteryTab.AutoScroll = true;
                toolTab.AutoScroll = true;
                officePostTab.AutoScroll = true;
                officeCommTab.AutoScroll = true;
                aboutTab.AutoScroll = true;
                wifiTab.AutoScrollMinSize = new Size(820, 700);
                batteryTab.AutoScrollMinSize = new Size(820, 700);
                toolTab.AutoScrollMinSize = new Size(860, 1260);
                officePostTab.AutoScrollMinSize = new Size(820, 1080);
                officeCommTab.AutoScrollMinSize = new Size(820, 700);
                aboutTab.AutoScrollMinSize = new Size(820, 860);
                tabs.TabPages.Add(wifiTab);
                tabs.TabPages.Add(batteryTab);
                tabs.TabPages.Add(toolTab);
                tabs.TabPages.Add(officePostTab);
                tabs.TabPages.Add(aboutTab);

                var listLabel = new Label { Text = "联网顺序", Location = new Point(18, 18), Size = new Size(100, 22) };
                var wifiList = new ListBox { Location = new Point(18, 46), Size = new Size(250, 478) };
                var editorLabel = new Label
                {
                    Text = "WiFi 编辑",
                    Location = new Point(292, 18),
                    Size = new Size(120, 22),
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point)
                };
                var ssidLabel = new Label { Text = "WiFi 名称", Location = new Point(292, 52), Size = new Size(120, 22) };
                var ssidBox = new TextBox { Location = new Point(292, 78), Size = new Size(370, 28), BorderStyle = BorderStyle.FixedSingle };
                var pwdLabel = new Label { Text = "WiFi 密码", Location = new Point(292, 118), Size = new Size(120, 22) };
                var pwdBox = new TextBox { Location = new Point(292, 144), Size = new Size(370, 28), UseSystemPasswordChar = true, BorderStyle = BorderStyle.FixedSingle };
                var authLabel = new Label { Text = "安全类型", Location = new Point(292, 184), Size = new Size(120, 22) };
                var authBox = new ComboBox
                {
                    Location = new Point(292, 210),
                    Size = new Size(170, 28),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                authBox.Items.AddRange(new object[] { "开放网络", "WPA2-PSK", "WPA-PSK", "WPA3-SAE" });
                authBox.SelectedIndex = 1;
                var encryptionLabel = new Label { Text = "加密方式", Location = new Point(492, 184), Size = new Size(120, 22) };
                var encryptionBox = new ComboBox
                {
                    Location = new Point(492, 210),
                    Size = new Size(170, 28),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                encryptionBox.Items.AddRange(new object[] { "AES", "TKIP", "无加密" });
                encryptionBox.SelectedIndex = 0;
                var tipLabel = new TextBox
                {
                    Text = "联网按钮会按左侧从上到下的顺序依次尝试。\r\n开放网络可不填密码。",
                    Location = new Point(292, 252),
                    Size = new Size(370, 44),
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    Multiline = true,
                    BackColor = wifiTab.BackColor,
                    ForeColor = Color.DimGray,
                    TabStop = false
                };
                var addButton = new Button { Text = "新增/更新", Location = new Point(292, 314), Size = new Size(100, 34) };
                var removeButton = new Button { Text = "删除", Location = new Point(404, 314), Size = new Size(78, 34) };
                var upButton = new Button { Text = "上移", Location = new Point(494, 314), Size = new Size(78, 34) };
                var downButton = new Button { Text = "下移", Location = new Point(584, 314), Size = new Size(78, 34) };
                wifiTab.Controls.AddRange(new Control[]
                {
                    listLabel, wifiList, editorLabel, ssidLabel, ssidBox, pwdLabel, pwdBox, authLabel, authBox, encryptionLabel, encryptionBox, tipLabel, addButton, removeButton, upButton, downButton
                });

                var batteryRuleLabel = new Label
                {
                    Text = "电池按钮规则",
                    Location = new Point(24, 22),
                    Size = new Size(120, 22),
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point)
                };
                var batteryRuleTip = new TextBox
                {
                    Text = "满足绿线显示绿色，小于红线显示红色。\r\n填 0 表示不启用该项规则。\r\n电池右击可手动标绿，再右击绿色按钮恢复真实检测。",
                    Location = new Point(24, 52),
                    Size = new Size(460, 72),
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    Multiline = true,
                    BackColor = batteryTab.BackColor,
                    ForeColor = Color.DimGray,
                    TabStop = false
                };
                var greenHealthLabel = new Label { Text = "健康度绿线(%)", Location = new Point(24, 138), Size = new Size(140, 22) };
                var greenHealthBox = new NumericUpDown { Location = new Point(24, 166), Size = new Size(150, 28), Minimum = 0, Maximum = 100, Value = _settings.BatteryGreenHealthPercent };
                var redHealthLabel = new Label { Text = "健康度红线(%)", Location = new Point(220, 138), Size = new Size(140, 22) };
                var redHealthBox = new NumericUpDown { Location = new Point(220, 166), Size = new Size(150, 28), Minimum = 0, Maximum = 100, Value = _settings.BatteryRedHealthPercent };
                var greenCapacityLabel = new Label { Text = "容量绿线(mWh)", Location = new Point(24, 226), Size = new Size(140, 22) };
                var greenCapacityBox = new NumericUpDown { Location = new Point(24, 254), Size = new Size(150, 28), Minimum = 0, Maximum = 200000, Increment = 1000, ThousandsSeparator = true, Value = _settings.BatteryGreenCapacityMWh };
                var redCapacityLabel = new Label { Text = "容量红线(mWh)", Location = new Point(220, 226), Size = new Size(140, 22) };
                var redCapacityBox = new NumericUpDown { Location = new Point(220, 254), Size = new Size(150, 28), Minimum = 0, Maximum = 200000, Increment = 1000, ThousandsSeparator = true, Value = _settings.BatteryRedCapacityMWh };
                batteryTab.Controls.AddRange(new Control[]
                {
                    batteryRuleLabel, batteryRuleTip, greenHealthLabel, greenHealthBox, redHealthLabel, redHealthBox, greenCapacityLabel, greenCapacityBox, redCapacityLabel, redCapacityBox
                });

                var toolConfigLabel = new Label
                {
                    Text = "工具路径配置",
                    Location = new Point(24, 22),
                    Size = new Size(120, 22),
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point)
                };
                var toolConfigTip = new TextBox
                {
                    Text = "Office 可设密钥源文件、剪切/复制模式、目标文件和默认打开软件。\r\n摄像头路径可填 app URI 或 exe；其他工具支持放在软件目录或 U 盘相对路径。",
                    Location = new Point(24, 52),
                    Size = new Size(680, 52),
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    Multiline = true,
                    BackColor = toolTab.BackColor,
                    ForeColor = Color.DimGray,
                    TabStop = false
                };
                var officeGroup = new GroupBox
                {
                    Text = "Office",
                    Location = new Point(24, 118),
                    Size = new Size(680, 492)
                };
                var officeSourceLabel = new Label { Text = "密钥文件路径", Location = new Point(18, 30), Size = new Size(140, 22) };
                var officeSourceBox = new TextBox { Location = new Point(18, 56), Size = new Size(642, 28), Text = _settings.OfficeKeySourcePath };
                var officeOperationLabel = new Label { Text = "操作方式", Location = new Point(18, 96), Size = new Size(120, 22) };
                var officeOperationBox = new ComboBox
                {
                    Location = new Point(18, 122),
                    Size = new Size(180, 28),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                officeOperationBox.Items.AddRange(new object[] { "复制密钥文件首行", "剪切密钥文件首行", "删除密钥文件首行" });
                officeOperationBox.SelectedItem = MapOfficeActionToDisplay(_settings.OfficeKeyOperation);
                var officeTargetLabel = new Label { Text = "目标文件路径(可多行)", Location = new Point(18, 160), Size = new Size(180, 22) };
                var officeTargetBox = new TextBox
                {
                    Location = new Point(18, 186),
                    Size = new Size(642, 58),
                    Text = BuildOfficeTargetPathsText(),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                officeTargetBox.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || !e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    var selectionStart = officeTargetBox.SelectionStart;
                    var newLine = Environment.NewLine;
                    officeTargetBox.SelectedText = newLine;
                    officeTargetBox.SelectionStart = selectionStart + newLine.Length;
                };
                var officeTargetTip = new Label
                {
                    Text = "可一行一个路径；程序会优先在这些路径里找已有文件，找不到再回退到当前桌面。任一设置处回车保存全部；Shift+回车换行。",
                    Location = new Point(18, 248),
                    Size = new Size(642, 54),
                    ForeColor = Color.DimGray
                };
                var officeLowKeyLabel = new Label { Text = "密钥不足提醒(<个)", Location = new Point(18, 304), Size = new Size(160, 22) };
                var officeLowKeyBox = new NumericUpDown { Location = new Point(180, 301), Size = new Size(120, 28), Minimum = 0, Maximum = 9999, Value = _settings.OfficeLowKeyWarningThreshold };
                var officeLowKeyTip = new Label { Text = "填 0 不提示", Location = new Point(314, 304), Size = new Size(120, 22), ForeColor = Color.DimGray };
                var officeAppLabel = new Label { Text = "Office 软件路径(可多行,按顺序尝试)", Location = new Point(18, 338), Size = new Size(240, 22) };
                var officeAppBox = new TextBox
                {
                    Location = new Point(18, 364),
                    Size = new Size(642, 52),
                    Text = BuildOfficeAppPathsText(),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                officeAppBox.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || !e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    var selectionStart = officeAppBox.SelectionStart;
                    var newLine = Environment.NewLine;
                    officeAppBox.SelectedText = newLine;
                    officeAppBox.SelectionStart = selectionStart + newLine.Length;
                };
                var officeAppTip = new Label
                {
                    Text = "可多行填写，程序会按顺序尝试打开；如果 Office 激活打开不对或路径失效，可在设置中修改。回车保存全部，Shift+回车换行。",
                    Location = new Point(18, 422),
                    Size = new Size(642, 62),
                    ForeColor = Color.DimGray
                };
                var officeCommUrlLabel = new Label { Text = "通信测试 URL", Location = new Point(18, 30), Size = new Size(140, 22) };
                var officeCommUrlBox = new TextBox
                {
                    Location = new Point(18, 56),
                    Size = new Size(642, 28),
                    Text = string.IsNullOrWhiteSpace(_settings.OfficeCommunicationTestUrl) ? DefaultOfficeCommunicationTestUrl : _settings.OfficeCommunicationTestUrl
                };
                var officeCommPayloadLabel = new Label { Text = "通信测试模拟文本", Location = new Point(18, 94), Size = new Size(160, 22) };
                var officeCommPayloadBox = new TextBox
                {
                    Location = new Point(18, 120),
                    Size = new Size(642, 56),
                    Text = string.IsNullOrWhiteSpace(_settings.OfficeCommunicationTestPayload) ? DefaultOfficeCommunicationTestPayload : _settings.OfficeCommunicationTestPayload,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                officeCommPayloadBox.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || !e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    var selectionStart = officeCommPayloadBox.SelectionStart;
                    var newLine = Environment.NewLine;
                    officeCommPayloadBox.SelectedText = newLine;
                    officeCommPayloadBox.SelectionStart = selectionStart + newLine.Length;
                };
                officeGroup.Controls.AddRange(new Control[]
                {
                    officeSourceLabel, officeSourceBox, officeOperationLabel, officeOperationBox, officeTargetLabel, officeTargetBox, officeTargetTip, officeLowKeyLabel, officeLowKeyBox, officeLowKeyTip, officeAppLabel, officeAppBox, officeAppTip
                });

                var toolGroup = new GroupBox
                {
                    Text = "其他检测工具",
                    Location = new Point(24, 626),
                    Size = new Size(680, 372)
                };
                var cameraPathLabel = new Label { Text = "摄像头检查路径/命令", Location = new Point(18, 30), Size = new Size(160, 22) };
                var cameraPathBox = new TextBox { Location = new Point(18, 56), Size = new Size(642, 28), Text = _settings.CameraCheckPath };
                var keyboardPathLabel = new Label { Text = "键盘检查路径", Location = new Point(18, 92), Size = new Size(140, 22) };
                var keyboardPathBox = new TextBox { Location = new Point(18, 118), Size = new Size(642, 28), Text = _settings.KeyboardCheckPath };
                var speakerPathLabel = new Label { Text = "扬声器测试路径", Location = new Point(18, 154), Size = new Size(140, 22) };
                var speakerPathBox = new TextBox { Location = new Point(160, 151), Size = new Size(500, 28), Text = _settings.SpeakerTestPath };
                var iconPathLabel = new Label { Text = "图标文件路径", Location = new Point(18, 194), Size = new Size(140, 22) };
                var iconPathBox = new TextBox { Location = new Point(160, 191), Size = new Size(500, 28), Text = _settings.IconFilePath };
                var windowsActivationTemplateLabel = new Label { Text = "Windows 激活复制模板", Location = new Point(18, 232), Size = new Size(180, 22) };
                var windowsActivationTemplateBox = new TextBox
                {
                    Location = new Point(18, 258),
                    Size = new Size(642, 52),
                    Text = BuildWindowsActivationCopyTemplateText(),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                windowsActivationTemplateBox.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || !e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    var selectionStart = windowsActivationTemplateBox.SelectionStart;
                    var newLine = Environment.NewLine;
                    windowsActivationTemplateBox.SelectedText = newLine;
                    windowsActivationTemplateBox.SelectionStart = selectionStart + newLine.Length;
                };
                var windowsActivationTemplateTip = new Label
                {
                    Text = "支持在这里自定义复制给用户的内容；可使用 {key} 代表输入的 Windows 密钥。回车保存全部，Shift+回车换行。",
                    Location = new Point(18, 314),
                    Size = new Size(642, 56),
                    ForeColor = Color.DimGray
                };
                toolGroup.Controls.AddRange(new Control[]
                {
                    cameraPathLabel, cameraPathBox, keyboardPathLabel, keyboardPathBox, speakerPathLabel, speakerPathBox, iconPathLabel, iconPathBox, windowsActivationTemplateLabel, windowsActivationTemplateBox, windowsActivationTemplateTip
                });

                var cleanupGroup = new GroupBox
                {
                    Text = "记录擦除附加清理",
                    Location = new Point(24, 1012),
                    Size = new Size(760, 228)
                };
                var cleanupPathLabel = new Label { Text = "文件/文件夹路径(可多行)", Location = new Point(18, 30), Size = new Size(220, 22) };
                var cleanupPathBox = new TextBox
                {
                    Location = new Point(18, 58),
                    Size = new Size(716, 96),
                    Text = BuildExtraCleanupPathsText(),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                cleanupPathBox.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || !e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    var selectionStart = cleanupPathBox.SelectionStart;
                    var newLine = Environment.NewLine;
                    cleanupPathBox.SelectedText = newLine;
                    cleanupPathBox.SelectionStart = selectionStart + newLine.Length;
                };
                var cleanupTipLabel = new Label
                {
                    Text = "一行一个路径。点“记录擦除”时会先删除这些文件/文件夹，最后再清一次回收站，尽量避免残留在垃圾箱里。",
                    Location = new Point(18, 162),
                    Size = new Size(716, 44),
                    ForeColor = Color.DimGray
                };
                cleanupGroup.Controls.AddRange(new Control[]
                {
                    cleanupPathLabel, cleanupPathBox, cleanupTipLabel
                });

                toolTab.Controls.AddRange(new Control[]
                {
                    toolConfigLabel, toolConfigTip, officeGroup, toolGroup, cleanupGroup
                });

                var officeCommTitleLabel = new Label
                {
                    Text = "Office 通信测试",
                    Location = new Point(24, 22),
                    Size = new Size(160, 22),
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point)
                };
                var officeCommTabTip = new TextBox
                {
                    Text = "这里配置 Office 的通信测试窗。右击 Office 按钮可打开测试窗；会先测试 URL 连通，再发送模拟文本并展示返回内容。\r\n全程不碰激活数据、不自动代填。",
                    Location = new Point(24, 52),
                    Size = new Size(680, 52),
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    Multiline = true,
                    BackColor = officeCommTab.BackColor,
                    ForeColor = Color.DimGray,
                    TabStop = false
                };
                var officeCommGroup = new GroupBox
                {
                    Text = "通信测试配置",
                    Location = new Point(24, 118),
                    Size = new Size(680, 250)
                };
                var officeCommTip = new TextBox
                {
                    Text = "测试窗会按顺序做：URL 连通测试 -> 发送模拟文本 -> 接收返回 -> 展示可复制内容。按 Esc 可直接关闭测试窗。",
                    Location = new Point(18, 182),
                    Size = new Size(642, 44),
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    Multiline = true,
                    BackColor = officeCommTab.BackColor,
                    ForeColor = Color.DimGray,
                    TabStop = false
                };
                officeCommGroup.Controls.AddRange(new Control[]
                {
                    officeCommUrlLabel, officeCommUrlBox, officeCommPayloadLabel, officeCommPayloadBox, officeCommTip
                });
                officeCommTab.Controls.AddRange(new Control[]
                {
                    officeCommTitleLabel, officeCommTabTip, officeCommGroup
                });

                var officePostTitleLabel = new Label
                {
                    Text = "Office 右键按键串",
                    Location = new Point(24, 22),
                    Size = new Size(160, 22),
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point)
                };
                var officePostTabTip = new TextBox
                {
                    Text = "这里专门设置 Office 右键菜单里的发送按键串功能。右击 Office 时会把焦点切到 Office 主窗口后发送；任一设置处回车保存全部，多行输入框里 Shift+回车换行。",
                    Location = new Point(24, 52),
                    Size = new Size(680, 52),
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    Multiline = true,
                    BackColor = officePostTab.BackColor,
                    ForeColor = Color.DimGray,
                    TabStop = false
                };
                var officePostGroup = new GroupBox
                {
                    Text = "右键发送设置",
                    Location = new Point(24, 118),
                    Size = new Size(680, 468)
                };
                var officePostActionEnabledBox = new CheckBox
                {
                    Text = "显示 Office 右键发送按键串",
                    Location = new Point(18, 32),
                    Size = new Size(240, 24),
                    Checked = _settings.OfficePostActionEnabled
                };
                var officePostConditionLabel = new Label { Text = "右键菜单文字", Location = new Point(18, 72), Size = new Size(120, 22) };
                var officePostConditionBox = new TextBox
                {
                    Location = new Point(18, 98),
                    Size = new Size(642, 68),
                    Text = GetOfficeSequenceMenuText(),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                officePostConditionBox.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || !e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    var selectionStart = officePostConditionBox.SelectionStart;
                    var newLine = Environment.NewLine;
                    officePostConditionBox.SelectedText = newLine;
                    officePostConditionBox.SelectionStart = selectionStart + newLine.Length;
                };
                var officePostConditionTip = new Label
                {
                    Text = "这里写右击 Office 时显示的菜单文字，例如：发送按键串 / 快速发送 / 输入产品密钥。",
                    Location = new Point(18, 172),
                    Size = new Size(642, 34),
                    ForeColor = Color.DimGray
                };
                var officePostSequenceLabel = new Label { Text = "发送按键串", Location = new Point(18, 214), Size = new Size(120, 22) };
                var officePostSequenceBox = new TextBox
                {
                    Location = new Point(18, 240),
                    Size = new Size(642, 140),
                    Text = string.IsNullOrWhiteSpace(_settings.OfficePostActionSequence) ? DefaultOfficePostActionSequence : _settings.OfficePostActionSequence,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                officePostSequenceBox.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || !e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    var selectionStart = officePostSequenceBox.SelectionStart;
                    var newLine = Environment.NewLine;
                    officePostSequenceBox.SelectedText = newLine;
                    officePostSequenceBox.SelectionStart = selectionStart + newLine.Length;
                };
                var officePostActionTip = new Label
                {
                    Text = "推荐每行带类型：按键:/文本:/停顿:。右击 Office 时会按顺序逐行发送到 Office 窗口。",
                    Location = new Point(18, 388),
                    Size = new Size(642, 52),
                    ForeColor = Color.DimGray
                };
                Action<string> appendOfficePostSequenceEntry = entry =>
                {
                    var normalizedEntry = (entry ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(normalizedEntry))
                    {
                        return;
                    }

                    var currentText = officePostSequenceBox.Text ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(currentText))
                    {
                        officePostSequenceBox.Text = normalizedEntry;
                    }
                    else
                    {
                        officePostSequenceBox.Text = currentText.TrimEnd() + Environment.NewLine + normalizedEntry;
                    }

                    officePostSequenceBox.SelectionStart = officePostSequenceBox.TextLength;
                    officePostSequenceBox.ScrollToCaret();
                };
                Action<string> appendOfficePostSequenceKeyToken = token =>
                {
                    foreach (var keyToken in SplitOfficeSequenceParts(token))
                    {
                        appendOfficePostSequenceEntry("按键:" + keyToken);
                    }
                };
                officePostGroup.Controls.AddRange(new Control[]
                {
                    officePostActionEnabledBox, officePostConditionLabel, officePostConditionBox, officePostConditionTip, officePostSequenceLabel, officePostSequenceBox, officePostActionTip
                });
                var officePostBuilderGroup = new GroupBox
                {
                    Text = "辅助生成",
                    Location = new Point(24, 600),
                    Size = new Size(680, 210)
                };
                string pendingOfficePostModifierToken = null;
                var officePostCaptureLabel = new Label { Text = "发送按键(点下面后直接按)", Location = new Point(18, 30), Size = new Size(220, 22) };
                var officePostCaptureBox = new TextBox
                {
                    Location = new Point(18, 56),
                    Size = new Size(642, 28),
                    ReadOnly = true,
                    TabStop = true,
                    Text = "点击这里后直接按键，程序会自动按“按键:xxx”追加到上面的规则里；单键和组合键都支持"
                };
                officePostCaptureBox.Tag = "office-post-capture";
                officePostCaptureBox.GotFocus += (_, __) =>
                {
                    officePostCaptureBox.SelectAll();
                };
                officePostCaptureBox.PreviewKeyDown += (_, e) =>
                {
                    if (e != null)
                    {
                        e.IsInputKey = true;
                    }
                };
                officePostCaptureBox.KeyDown += (_, e) =>
                {
                    string modifierToken;
                    if (TryBuildStandaloneOfficeModifierToken(e, out modifierToken))
                    {
                        pendingOfficePostModifierToken = modifierToken;
                        officePostCaptureBox.Text = "松开后追加: " + modifierToken;
                        officePostCaptureBox.SelectAll();
                        if (e != null)
                        {
                            e.SuppressKeyPress = true;
                            e.Handled = true;
                        }
                        return;
                    }

                    pendingOfficePostModifierToken = null;
                    var token = BuildOfficePostTokenFromKeyEvent(e);
                    if (!string.IsNullOrWhiteSpace(token))
                    {
                        appendOfficePostSequenceKeyToken(token);
                        officePostCaptureBox.Text = "已追加: " + token;
                        officePostCaptureBox.SelectAll();
                    }

                    if (e != null)
                    {
                        e.SuppressKeyPress = true;
                        e.Handled = true;
                    }
                };
                officePostCaptureBox.KeyUp += (_, e) =>
                {
                    string modifierToken;
                    if (!TryBuildStandaloneOfficeModifierToken(e, out modifierToken))
                    {
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(pendingOfficePostModifierToken)
                        || !string.Equals(pendingOfficePostModifierToken, modifierToken, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    appendOfficePostSequenceKeyToken(modifierToken);
                    officePostCaptureBox.Text = "已追加: " + modifierToken;
                    officePostCaptureBox.SelectAll();
                    pendingOfficePostModifierToken = null;
                    if (e != null)
                    {
                        e.SuppressKeyPress = true;
                        e.Handled = true;
                    }
                };
                var officePostTextLabel = new Label { Text = "发送文本(可直接编写/粘贴)", Location = new Point(18, 94), Size = new Size(220, 22) };
                var officePostTextBox = new TextBox
                {
                    Location = new Point(18, 120),
                    Size = new Size(516, 28)
                };
                var officePostAddTextButton = new Button
                {
                    Text = "追加文本",
                    Location = new Point(548, 118),
                    Size = new Size(112, 32)
                };
                officePostAddTextButton.Click += (_, __) =>
                {
                    var literalText = officePostTextBox.Text ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(literalText))
                    {
                        return;
                    }

                    appendOfficePostSequenceEntry("文本:" + literalText.Replace("\r", " ").Replace("\n", " "));
                    officePostTextBox.Clear();
                    officePostTextBox.Focus();
                };
                var officePostBuilderTip = new Label
                {
                    Text = "“发送按键”“追加文本”“停顿”都会按带类型的一行写入上面的“发送按键串”，方便直接阅读和维护。",
                    Location = new Point(18, 190),
                    Size = new Size(642, 18),
                    ForeColor = Color.DimGray
                };
                var officePostDelay200Button = new Button { Text = "停顿200ms", Location = new Point(18, 160), Size = new Size(104, 32) };
                var officePostDelay500Button = new Button { Text = "停顿500ms", Location = new Point(130, 160), Size = new Size(104, 32) };
                var officePostDelay1000Button = new Button { Text = "停顿1000ms", Location = new Point(242, 160), Size = new Size(116, 32) };
                var officePostClearButton = new Button { Text = "清空按键串", Location = new Point(474, 160), Size = new Size(90, 32) };
                var officePostDefaultButton = new Button { Text = "恢复默认", Location = new Point(570, 160), Size = new Size(90, 32) };
                officePostDelay200Button.Click += (_, __) => appendOfficePostSequenceEntry("停顿:200");
                officePostDelay500Button.Click += (_, __) => appendOfficePostSequenceEntry("停顿:500");
                officePostDelay1000Button.Click += (_, __) => appendOfficePostSequenceEntry("停顿:1000");
                officePostClearButton.Click += (_, __) => officePostSequenceBox.Clear();
                officePostDefaultButton.Click += (_, __) => officePostSequenceBox.Text = DefaultOfficePostActionSequence;
                officePostBuilderGroup.Controls.AddRange(new Control[]
                {
                    officePostCaptureLabel, officePostCaptureBox, officePostTextLabel, officePostTextBox, officePostAddTextButton,
                    officePostDelay200Button, officePostDelay500Button, officePostDelay1000Button, officePostClearButton, officePostDefaultButton, officePostBuilderTip
                });
                var officePostAhkGroup = new GroupBox
                {
                    Text = "AHK 嵌入代码",
                    Location = new Point(24, 824),
                    Size = new Size(680, 238)
                };
                var officePostAhkEnabledBox = new CheckBox
                {
                    Text = "命中规则后额外执行 AHK 代码",
                    Location = new Point(18, 30),
                    Size = new Size(250, 24),
                    Checked = _settings.OfficePostAhkEnabled
                };
                var officePostAhkPathLabel = new Label { Text = "AHK 程序路径", Location = new Point(18, 64), Size = new Size(120, 22) };
                var officePostAhkPathBox = new TextBox
                {
                    Location = new Point(18, 90),
                    Size = new Size(642, 28),
                    Text = _settings.OfficePostAhkPath
                };
                var officePostAhkScriptLabel = new Label { Text = "AHK 代码", Location = new Point(18, 126), Size = new Size(120, 22) };
                var officePostAhkScriptBox = new TextBox
                {
                    Location = new Point(18, 152),
                    Size = new Size(642, 52),
                    Text = string.IsNullOrWhiteSpace(_settings.OfficePostAhkScript) ? string.Empty : _settings.OfficePostAhkScript,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                officePostAhkScriptBox.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || !e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    var selectionStart = officePostAhkScriptBox.SelectionStart;
                    var newLine = Environment.NewLine;
                    officePostAhkScriptBox.SelectedText = newLine;
                    officePostAhkScriptBox.SelectionStart = selectionStart + newLine.Length;
                };
                var officePostAhkTip = new Label
                {
                    Text = "填写 AutoHotkey.exe 路径；命中规则后会额外执行这里的 AHK 代码。支持占位符：{key} {condition} {appdir}。",
                    Location = new Point(18, 208),
                    Size = new Size(642, 22),
                    ForeColor = Color.DimGray
                };
                officePostAhkGroup.Controls.AddRange(new Control[]
                {
                    officePostAhkEnabledBox, officePostAhkPathLabel, officePostAhkPathBox, officePostAhkScriptLabel, officePostAhkScriptBox, officePostAhkTip
                });
                officePostTab.Controls.AddRange(new Control[]
                {
                    officePostTitleLabel, officePostTabTip, officePostGroup, officePostBuilderGroup
                });

                var aboutTitleLabel = new Label
                {
                    Text = "关于 " + GetConfiguredAppTitle(),
                    Location = new Point(24, 22),
                    Size = new Size(140, 22),
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point)
                };
                var aboutTipLabel = new Label
                {
                    Text = "作者、软件名、更新时间与链接信息可直接复制；修改软件名/更新时间/GitHub/支持作者链接需口令。",
                    Location = new Point(24, 56),
                    Size = new Size(680, 36),
                    ForeColor = Color.DimGray
                };
                var aboutCopyButton = new Button
                {
                    Text = "复制全部信息",
                    Location = new Point(584, 18),
                    Size = new Size(120, 32)
                };
                var aboutNameLabel = new Label { Text = "软件名称", Location = new Point(24, 108), Size = new Size(120, 22) };
                var aboutNameBox = new TextBox { Location = new Point(24, 134), Size = new Size(220, 28), Text = GetConfiguredAppName() };
                var aboutUpdatedAtLabel = new Label { Text = "软件更新时间", Location = new Point(270, 108), Size = new Size(120, 22) };
                var aboutUpdatedAtBox = new TextBox { Location = new Point(270, 134), Size = new Size(220, 28), Text = string.IsNullOrWhiteSpace(_settings.SoftwareUpdatedAt) ? DefaultSoftwareUpdatedAt : _settings.SoftwareUpdatedAt };
                var aboutGithubLabel = new Label { Text = "GitHub 链接", Location = new Point(24, 174), Size = new Size(120, 22) };
                var aboutGithubBox = new TextBox { Location = new Point(24, 200), Size = new Size(520, 28), Text = _settings.GithubUrl };
                var aboutGithubButton = new Button { Text = "打开 GitHub", Location = new Point(560, 198), Size = new Size(144, 32) };
                var aboutDonateLabel = new Label { Text = "支持作者链接", Location = new Point(24, 240), Size = new Size(120, 22) };
                var aboutDonateBox = new TextBox { Location = new Point(24, 266), Size = new Size(520, 28), Text = _settings.DonateUrl };
                var aboutDonateButton = new Button { Text = "打开支持作者链接", Location = new Point(560, 264), Size = new Size(144, 32) };
                var aboutTextBox = new TextBox
                {
                    Location = new Point(24, 318),
                    Size = new Size(680, 308),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    ShortcutsEnabled = true,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point)
                };
                Action refreshAboutPreview = () =>
                {
                    aboutTextBox.Text = BuildAboutText(aboutNameBox.Text, aboutUpdatedAtBox.Text, aboutGithubBox.Text, aboutDonateBox.Text);
                };
                aboutNameBox.TextChanged += (_, __) => refreshAboutPreview();
                aboutUpdatedAtBox.TextChanged += (_, __) => refreshAboutPreview();
                aboutGithubBox.TextChanged += (_, __) => refreshAboutPreview();
                aboutDonateBox.TextChanged += (_, __) => refreshAboutPreview();
                aboutGithubButton.Click += (_, __) => OpenConfiguredWebLink(aboutGithubBox.Text, "GitHub");
                aboutDonateButton.Click += (_, __) => OpenConfiguredWebLink(aboutDonateBox.Text, "支持作者");
                aboutCopyButton.Click += (_, __) =>
                {
                    try
                    {
                        Clipboard.SetText(aboutTextBox.Text ?? string.Empty);
                    }
                    catch
                    {
                        ShowErrorMessage("复制失败，请稍后重试。", "关于");
                    }
                };
                refreshAboutPreview();
                aboutTab.Controls.AddRange(new Control[]
                {
                    aboutTitleLabel, aboutTipLabel, aboutCopyButton,
                    aboutNameLabel, aboutNameBox, aboutUpdatedAtLabel, aboutUpdatedAtBox,
                    aboutGithubLabel, aboutGithubBox, aboutGithubButton,
                    aboutDonateLabel, aboutDonateBox, aboutDonateButton,
                    aboutTextBox
                });

                var okButton = new Button { Text = "保存设置", Location = new Point(608, 760), Size = new Size(92, 36) };
                var cancelButton = new Button { Text = "取消", Location = new Point(710, 760), Size = new Size(74, 36) };

                Action refreshWifiList = () =>
                {
                    wifiList.Items.Clear();
                    for (var i = 0; i < _settings.WifiProfiles.Count; i++)
                    {
                        var profile = _settings.WifiProfiles[i];
                        var suffix = string.Equals(profile.Authentication, "open", StringComparison.OrdinalIgnoreCase)
                            ? "开放"
                            : MapAuthenticationToDisplay(profile.Authentication);
                        wifiList.Items.Add(string.Format("{0}. {1} [{2}]", i + 1, profile.Ssid, suffix));
                    }
                };

                Action<int> loadProfileToEditor = index =>
                {
                    if (index < 0 || index >= _settings.WifiProfiles.Count)
                    {
                        return;
                    }

                    var profile = _settings.WifiProfiles[index];
                    ssidBox.Text = profile.Ssid;
                    pwdBox.Text = profile.Password;
                    SelectAuthDisplayValue(authBox, profile.Authentication);
                    SelectEncryptionDisplayValue(encryptionBox, profile.Encryption);
                };

                authBox.SelectedIndexChanged += (_, __) =>
                {
                    var isOpen = authBox.SelectedItem != null && authBox.SelectedItem.ToString() == "开放网络";
                    pwdBox.Enabled = !isOpen;
                    pwdBox.UseSystemPasswordChar = !isOpen;
                    if (isOpen)
                    {
                        pwdBox.Text = string.Empty;
                        encryptionBox.SelectedItem = "无加密";
                        encryptionBox.Enabled = false;
                    }
                    else
                    {
                        if (encryptionBox.SelectedItem == null || encryptionBox.SelectedItem.ToString() == "无加密")
                        {
                            encryptionBox.SelectedItem = "AES";
                        }
                        encryptionBox.Enabled = true;
                    }
                };

                wifiList.SelectedIndexChanged += (_, __) => loadProfileToEditor(wifiList.SelectedIndex);

                addButton.Click += (_, __) =>
                {
                    if (string.IsNullOrWhiteSpace(ssidBox.Text))
                    {
                        ShowErrorMessage("请输入 WiFi 名称。", "设置");
                        tabs.SelectedTab = wifiTab;
                        return;
                    }

                    var authentication = MapDisplayToAuthentication(authBox.SelectedItem == null ? string.Empty : authBox.SelectedItem.ToString());
                    var encryption = string.Equals(authentication, "open", StringComparison.OrdinalIgnoreCase)
                        ? "none"
                        : MapDisplayToEncryption(encryptionBox.SelectedItem == null ? "AES" : encryptionBox.SelectedItem.ToString());

                    if (!string.Equals(authentication, "open", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(pwdBox.Text))
                    {
                        ShowErrorMessage("当前安全类型需要输入 WiFi 密码。", "设置");
                        tabs.SelectedTab = wifiTab;
                        return;
                    }

                    var existing = _settings.WifiProfiles.FindIndex(p => string.Equals(p.Ssid, ssidBox.Text.Trim(), StringComparison.OrdinalIgnoreCase));
                    var profile = new WifiProfile
                    {
                        Ssid = ssidBox.Text.Trim(),
                        Password = pwdBox.Text,
                        Authentication = authentication,
                        Encryption = encryption
                    };

                    if (existing >= 0)
                    {
                        _settings.WifiProfiles[existing] = profile;
                        refreshWifiList();
                        wifiList.SelectedIndex = existing;
                    }
                    else
                    {
                        _settings.WifiProfiles.Add(profile);
                        refreshWifiList();
                        wifiList.SelectedIndex = _settings.WifiProfiles.Count - 1;
                    }
                };

                removeButton.Click += (_, __) =>
                {
                    if (wifiList.SelectedIndex < 0)
                    {
                        return;
                    }

                    _settings.WifiProfiles.RemoveAt(wifiList.SelectedIndex);
                    refreshWifiList();
                    ssidBox.Text = string.Empty;
                    pwdBox.Text = string.Empty;
                    SelectAuthDisplayValue(authBox, "WPA2PSK");
                    SelectEncryptionDisplayValue(encryptionBox, "AES");
                };

                upButton.Click += (_, __) =>
                {
                    var index = wifiList.SelectedIndex;
                    if (index <= 0)
                    {
                        return;
                    }

                    var item = _settings.WifiProfiles[index];
                    _settings.WifiProfiles.RemoveAt(index);
                    _settings.WifiProfiles.Insert(index - 1, item);
                    refreshWifiList();
                    wifiList.SelectedIndex = index - 1;
                };

                downButton.Click += (_, __) =>
                {
                    var index = wifiList.SelectedIndex;
                    if (index < 0 || index >= _settings.WifiProfiles.Count - 1)
                    {
                        return;
                    }

                    var item = _settings.WifiProfiles[index];
                    _settings.WifiProfiles.RemoveAt(index);
                    _settings.WifiProfiles.Insert(index + 1, item);
                    refreshWifiList();
                    wifiList.SelectedIndex = index + 1;
                };

                okButton.Click += (_, __) =>
                {
                    var batteryGreenHealth = (int)greenHealthBox.Value;
                    var batteryRedHealth = (int)redHealthBox.Value;
                    var batteryGreenCapacity = (int)greenCapacityBox.Value;
                    var batteryRedCapacity = (int)redCapacityBox.Value;

                    if (batteryGreenHealth > 0 && batteryRedHealth > 0 && batteryGreenHealth <= batteryRedHealth)
                    {
                        ShowErrorMessage("电池健康度绿线必须大于红线。", "设置");
                        tabs.SelectedTab = batteryTab;
                        return;
                    }

                    if (batteryGreenCapacity > 0 && batteryRedCapacity > 0 && batteryGreenCapacity <= batteryRedCapacity)
                    {
                        ShowErrorMessage("电池容量绿线必须大于红线。", "设置");
                        tabs.SelectedTab = batteryTab;
                        return;
                    }

                    var normalizedGithubUrl = aboutGithubBox.Text.Trim();
                    var normalizedDonateUrl = aboutDonateBox.Text.Trim();
                    var normalizedDisplayName = aboutNameBox.Text.Trim();
                    var normalizedUpdatedAt = string.IsNullOrWhiteSpace(aboutUpdatedAtBox.Text) ? DefaultSoftwareUpdatedAt : aboutUpdatedAtBox.Text.Trim();
                    var protectedFieldsChanged = !string.Equals(normalizedDisplayName, GetConfiguredAppName(), StringComparison.Ordinal)
                        || !string.Equals(normalizedUpdatedAt, (_settings.SoftwareUpdatedAt ?? string.Empty).Trim(), StringComparison.Ordinal)
                        || !string.Equals(normalizedGithubUrl, (_settings.GithubUrl ?? string.Empty).Trim(), StringComparison.Ordinal)
                        || !string.Equals(normalizedDonateUrl, (_settings.DonateUrl ?? string.Empty).Trim(), StringComparison.Ordinal);

                    if (protectedFieldsChanged && !PromptForLinkEditPassword())
                    {
                        tabs.SelectedTab = aboutTab;
                        return;
                    }

                    _settings.BatteryGreenHealthPercent = batteryGreenHealth;
                    _settings.BatteryRedHealthPercent = batteryRedHealth;
                    _settings.BatteryGreenCapacityMWh = batteryGreenCapacity;
                    _settings.BatteryRedCapacityMWh = batteryRedCapacity;
                    _settings.OfficeKeySourcePath = officeSourceBox.Text.Trim();
                    var officeTargetPaths = SplitMultiLinePaths(officeTargetBox.Text);
                    _settings.OfficeKeyTargetPaths.Clear();
                    _settings.OfficeKeyTargetPaths.AddRange(officeTargetPaths);
                    _settings.OfficeKeyTargetPath = officeTargetPaths.FirstOrDefault() ?? string.Empty;
                    _settings.OfficeKeyOperation = MapDisplayToOfficeAction(officeOperationBox.SelectedItem == null ? "复制密钥文件首行" : officeOperationBox.SelectedItem.ToString());
                    _settings.OfficeLowKeyWarningThreshold = (int)officeLowKeyBox.Value;
                    var officeAppPaths = SplitMultiLinePaths(officeAppBox.Text);
                    _settings.OfficeAppPaths.Clear();
                    _settings.OfficeAppPaths.AddRange(officeAppPaths);
                    _settings.OfficeAppPath = officeAppPaths.FirstOrDefault() ?? string.Empty;
                    _settings.OfficeCommunicationTestUrl = officeCommUrlBox.Text.Trim();
                    _settings.OfficeCommunicationTestPayload = officeCommPayloadBox.Text ?? string.Empty;
                _settings.OfficePostActionEnabled = officePostActionEnabledBox.Checked;
                _settings.OfficePostActionCondition = string.IsNullOrWhiteSpace(officePostConditionBox.Text) ? DefaultOfficeSequenceMenuText : officePostConditionBox.Text.Trim();
                _settings.OfficePostActionSequence = string.IsNullOrWhiteSpace(officePostSequenceBox.Text) ? DefaultOfficePostActionSequence : officePostSequenceBox.Text.Trim();
                _settings.OfficePostAhkEnabled = officePostAhkEnabledBox.Checked;
                _settings.OfficePostAhkPath = officePostAhkPathBox.Text.Trim();
                _settings.OfficePostAhkScript = officePostAhkScriptBox.Text ?? string.Empty;
                _settings.CameraCheckPath = cameraPathBox.Text.Trim();
                _settings.KeyboardCheckPath = keyboardPathBox.Text.Trim();
                _settings.SpeakerTestPath = speakerPathBox.Text.Trim();
                    _settings.IconFilePath = iconPathBox.Text.Trim();
                    _settings.WindowsActivationCopyTemplate = string.IsNullOrWhiteSpace(windowsActivationTemplateBox.Text) ? DefaultWindowsActivationCopyTemplate : windowsActivationTemplateBox.Text.Trim();
                    _settings.ExtraCleanupPaths.Clear();
                    _settings.ExtraCleanupPaths.AddRange(SplitMultiLinePaths(cleanupPathBox.Text));
                    _settings.SoftwareDisplayName = normalizedDisplayName;
                    _settings.SoftwareUpdatedAt = normalizedUpdatedAt;
                    _settings.GithubUrl = normalizedGithubUrl;
                    _settings.DonateUrl = normalizedDonateUrl;
                    _settings.Save();
                    RefreshOfficeSequenceMenuItem();
                    RefreshAppTitlePresentation();
                    RefreshDashboard();
                    dialog.DialogResult = DialogResult.OK;
                    dialog.Close();
                };

                cancelButton.Click += (_, __) =>
                {
                    dialog.DialogResult = DialogResult.Cancel;
                    dialog.Close();
                };

                dialog.Controls.Add(tabs);
                dialog.Controls.Add(okButton);
                dialog.Controls.Add(cancelButton);
                ApplyDialogDpiScaling(dialog, new Size(820, 820));

                dialog.KeyPreview = true;
                dialog.AcceptButton = okButton;
                dialog.CancelButton = cancelButton;
                dialog.KeyDown += (_, e) =>
                {
                    if (e == null || e.KeyCode != Keys.Enter || e.Alt)
                    {
                        return;
                    }

                    Control focused = dialog.ActiveControl;
                    while (focused is ContainerControl && ((ContainerControl)focused).ActiveControl != null)
                    {
                        focused = ((ContainerControl)focused).ActiveControl;
                    }

                    if (focused != null && string.Equals(Convert.ToString(focused.Tag), "office-post-capture", StringComparison.Ordinal))
                    {
                        return;
                    }

                    var textBox = focused as TextBox;
                    if (textBox != null && textBox.Multiline && e.Shift)
                    {
                        return;
                    }

                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    okButton.PerformClick();
                };
                refreshWifiList();
                if (_settings.WifiProfiles.Count > 0)
                {
                    wifiList.SelectedIndex = 0;
                }
                else
                {
                    SelectAuthDisplayValue(authBox, "WPA2PSK");
                    SelectEncryptionDisplayValue(encryptionBox, "AES");
                }
                dialog.Shown += (_, __) =>
                {
                    tabs.SelectedTab = wifiTab;
                    ssidBox.Focus();
                };
                dialog.ShowDialog(this);
            }
        }

        private static string MapDisplayToAuthentication(string display)
        {
            switch (display)
            {
                case "开放网络":
                    return "open";
                case "WPA-PSK":
                    return "WPAPSK";
                case "WPA3-SAE":
                    return "WPA3SAE";
                default:
                    return "WPA2PSK";
            }
        }

        private static string MapAuthenticationToDisplay(string value)
        {
            switch (value)
            {
                case "open":
                    return "开放网络";
                case "WPAPSK":
                    return "WPA-PSK";
                case "WPA3SAE":
                    return "WPA3-SAE";
                default:
                    return "WPA2-PSK";
            }
        }

        private static string MapDisplayToEncryption(string display)
        {
            return display == "无加密" ? "none" : display;
        }

        private static string MapEncryptionToDisplay(string value)
        {
            return value == "none" ? "无加密" : value;
        }

        private static string MapOfficeActionToDisplay(string value)
        {
            if (string.Equals(value, "cut", StringComparison.OrdinalIgnoreCase))
            {
                return "剪切密钥文件首行";
            }

            if (string.Equals(value, "delete", StringComparison.OrdinalIgnoreCase))
            {
                return "删除密钥文件首行";
            }

            return "复制密钥文件首行";
        }

        private static string MapDisplayToOfficeAction(string value)
        {
            if (value == "剪切密钥文件首行" || value == "剪切首行")
            {
                return "cut";
            }

            if (value == "删除密钥文件首行" || value == "删除首行")
            {
                return "delete";
            }

            return "copy";
        }

        private static string GetOfficeActionShortText(string value)
        {
            if (string.Equals(value, "cut", StringComparison.OrdinalIgnoreCase))
            {
                return "剪切";
            }

            if (string.Equals(value, "delete", StringComparison.OrdinalIgnoreCase))
            {
                return "删除";
            }

            return "复制";
        }

        private static void SelectAuthDisplayValue(ComboBox comboBox, string authentication)
        {
            comboBox.SelectedItem = MapAuthenticationToDisplay(authentication);
            if (comboBox.SelectedIndex < 0)
            {
                comboBox.SelectedItem = "WPA2-PSK";
            }
        }

        private static void SelectEncryptionDisplayValue(ComboBox comboBox, string encryption)
        {
            comboBox.SelectedItem = MapEncryptionToDisplay(encryption);
            if (comboBox.SelectedIndex < 0)
            {
                comboBox.SelectedItem = "AES";
            }
        }

        private void OpenWindowsActivation()
        {
            _lastWindowsTimeSyncStatus = TryResyncWindowsTime();
            UpdateWindowsActivationButtonPresentation();

            var opened = TryStartTarget("ms-settings:activation");
            if (!opened)
            {
                opened = TryStartTarget("slui.exe");
            }

            if (!opened)
            {
                StartTarget("control.exe", "/name Microsoft.System");
            }

            QueueWindowsActivationRefresh();
        }

        private void QueueWindowsActivationRefresh()
        {
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                foreach (var delay in new[] { 2200, 5200 })
                {
                    System.Threading.Thread.Sleep(delay);
                    bool activated;
                    string statusText;
                    TryGetWindowsActivationState(out activated, out statusText);
                    try
                    {
                        BeginInvoke((MethodInvoker)delegate
                        {
                            _lastWindowsActivationStatus = statusText;
                            ApplyWindowsActivationState(activated, statusText);
                        });
                    }
                    catch
                    {
                    }
                }
            });
        }

        private void OpenConfigurationCheck()
        {
            var aboutOpened = TryStartTarget("ms-settings:about");
            if (!aboutOpened)
            {
                aboutOpened = TryStartTarget("control.exe", "/name Microsoft.System");
            }

            if (!aboutOpened)
            {
                aboutOpened = StartTarget("sysdm.cpl");
            }
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                System.Threading.Thread.Sleep(450);
                try
                {
                    StartTarget("diskmgmt.msc");
                }
                catch
                {
                }
            });
        }

        private void OpenBatteryCheck()
        {
            var toolPath = FindBatteryToolPath();
            if (toolPath == null)
            {
                ShowErrorMessage("未找到 BatteryInfoView，请确认它放在软件目录、U 盘根目录，或图吧工具箱目录里。");
                return;
            }

            StartTarget(toolPath);
        }

        private void OpenCameraCheck()
        {
            var candidates = new List<string>
            {
                _settings.CameraCheckPath
            };
            candidates.AddRange(new[]
            {
                "microsoft.windows.camera:",
                @"C:\Program Files\Windows Camera\WindowsCamera.exe",
                @"C:\Program Files\WindowsApps\Microsoft.WindowsCamera_2022.2501.6.0_x64__8wekyb3d8bbwe\WindowsCamera.exe"
            });

            foreach (var candidate in candidates.Where(candidate => !string.IsNullOrWhiteSpace(candidate)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (IsShellUri(candidate))
                {
                    if (TryStartTarget(candidate))
                    {
                        return;
                    }
                }
                else if (File.Exists(candidate) && TryStartTarget(candidate))
                {
                        return;
                }
            }

            if (TryStartTarget("control.exe", "/name Microsoft.DeviceManager"))
            {
                return;
            }

            ShowErrorMessage("当前系统未找到可直接打开画面的相机程序。\r\n可在设置里配置“摄像头检查路径/命令”。");
        }

        private void OpenKeyboardCheck()
        {
            var toolPath = BuildPortableCandidates(
                _settings.KeyboardCheckPath,
                "Keyboard Test Utility.exe",
                "键盘测试.exe").FirstOrDefault(File.Exists);
            if (toolPath == null)
            {
                ShowErrorMessage("未找到键盘检测工具，请检查设置里的“键盘检查路径”，或确认默认工具存在。");
                return;
            }

            StartTarget(toolPath);
        }

        private static bool IsShellUri(string candidate)
        {
            return !string.IsNullOrWhiteSpace(candidate) && candidate.Contains(":") && !candidate.Contains(@":\");
        }

        private string GetApplicationDirectory()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            if (string.IsNullOrWhiteSpace(baseDir))
            {
                baseDir = Directory.GetCurrentDirectory();
            }

            return baseDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        private string GetApplicationDriveRoot()
        {
            return Path.GetPathRoot(GetApplicationDirectory()) ?? string.Empty;
        }

        private string ResolveConfiguredPath(string configuredPath)
        {
            var trimmed = (configuredPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return string.Empty;
            }

            if (IsShellUri(trimmed) || Path.IsPathRooted(trimmed))
            {
                return trimmed;
            }

            return Path.GetFullPath(Path.Combine(GetApplicationDirectory(), trimmed));
        }

        private List<string> BuildPortableCandidates(string configuredPath, params string[] relativeCandidates)
        {
            var result = new List<string>();
            Action<string> addCandidate = candidate =>
            {
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    return;
                }

                if (!result.Any(existing => string.Equals(existing, candidate, StringComparison.OrdinalIgnoreCase)))
                {
                    result.Add(candidate);
                }
            };

            addCandidate(ResolveConfiguredPath(configuredPath));

            var appDir = GetApplicationDirectory();
            var driveRoot = GetApplicationDriveRoot();
            foreach (var relativeCandidate in relativeCandidates)
            {
                if (string.IsNullOrWhiteSpace(relativeCandidate))
                {
                    continue;
                }

                if (Path.IsPathRooted(relativeCandidate))
                {
                    addCandidate(relativeCandidate);
                    continue;
                }

                addCandidate(Path.GetFullPath(Path.Combine(appDir, relativeCandidate)));
                if (!string.IsNullOrWhiteSpace(driveRoot))
                {
                    addCandidate(Path.GetFullPath(Path.Combine(driveRoot, relativeCandidate)));
                }
            }

            return result;
        }

        private string GetDefaultOfficeTargetPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), OfficeKeyTargetPath);
        }

        private string BuildOfficeTargetPathsText()
        {
            return string.Join("\r\n", GetConfiguredOfficeTargetPaths().Select(ResolveOfficeTargetPath).ToArray());
        }

        private string BuildExtraCleanupPathsText()
        {
            return string.Join(
                "\r\n",
                _settings.ExtraCleanupPaths
                    .Where(path => !string.IsNullOrWhiteSpace(path))
                    .Select(path => ResolveConfiguredPath(path.Trim()))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray());
        }

        private string BuildOfficeAppPathsText()
        {
            var paths = _settings.OfficeAppPaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (paths.Count == 0 && !string.IsNullOrWhiteSpace(_settings.OfficeAppPath))
            {
                paths.Add(_settings.OfficeAppPath.Trim());
            }

            return string.Join("\r\n", paths.ToArray());
        }

        private string BuildWindowsActivationCopyTemplateText()
        {
            var template = (_settings.WindowsActivationCopyTemplate ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(template) ? DefaultWindowsActivationCopyTemplate : template;
        }

        private string BuildWindowsActivationClipboardText(string rawKey)
        {
            var template = BuildWindowsActivationCopyTemplateText();
            var normalizedKey = string.IsNullOrWhiteSpace(rawKey) ? "你的密钥" : rawKey.Trim();
            if (template.IndexOf("{key}", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return Regex.Replace(template, "\\{key\\}", normalizedKey, RegexOptions.IgnoreCase).Trim();
            }

            return string.IsNullOrWhiteSpace(normalizedKey)
                ? template.Trim()
                : (template.TrimEnd() + " " + normalizedKey).Trim();
        }

        private string NormalizeConfiguredUrl(string value)
        {
            var trimmed = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return string.Empty;
            }

            return trimmed;
        }

        private string BuildAboutText(string softwareName, string softwareUpdatedAt, string githubUrl, string donateUrl)
        {
            var displayName = string.IsNullOrWhiteSpace(softwareName) ? GetConfiguredAppTitle() : softwareName.Trim();
            var updatedAtText = string.IsNullOrWhiteSpace(softwareUpdatedAt) ? DefaultSoftwareUpdatedAt : softwareUpdatedAt.Trim();
            var githubText = string.IsNullOrWhiteSpace(githubUrl) ? "未设置" : githubUrl.Trim();
            var donateText = string.IsNullOrWhiteSpace(donateUrl) ? "未设置" : donateUrl.Trim();
            return
                displayName + "\r\n" +
                "版本：" + AppVersion + "\r\n" +
                "软件更新时间：" + updatedAtText + "\r\n" +
                "\r\n" +
                "作者：保安大王\r\n" +
                "邮箱：wangkuan.math@gmail.com\r\n" +
                "GitHub：" + githubText + "\r\n" +
                "支持作者链接：" + donateText + "\r\n" +
                "用途：电脑交付前基础功能检测与状态核验\r\n" +
                "适用系统：Windows 10 / Windows 11\r\n" +
                "支持项目：联网、激活、配置、电池、摄像头、键盘、记录清理等\r\n" +
                "备注：部分检测功能需管理员权限或外部工具支持";
        }

        private void OpenConfiguredWebLink(string url, string title)
        {
            var normalizedUrl = NormalizeConfiguredUrl(url);
            if (string.IsNullOrWhiteSpace(normalizedUrl))
            {
                ShowErrorMessage("当前还没有设置链接。", title);
                return;
            }

            if (!Regex.IsMatch(normalizedUrl, @"^[a-z][a-z0-9+\-.]*://", RegexOptions.IgnoreCase))
            {
                normalizedUrl = "https://" + normalizedUrl;
            }

            if (!TryStartTarget(normalizedUrl))
            {
                ShowErrorMessage("链接打开失败，请确认设置里的地址是否正确。", title);
            }
        }

        private bool PromptForLinkEditPassword()
        {
            using (var dialog = new Form())
            {
                dialog.Text = "验证口令";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.ClientSize = new Size(360, 152);
                dialog.Font = Font;

                var label = new Label
                {
                    Text = "修改 GitHub/支持作者链接需要口令。",
                    Location = new Point(18, 18),
                    Size = new Size(320, 22)
                };
                var passwordBox = new TextBox
                {
                    Location = new Point(18, 54),
                    Size = new Size(320, 28),
                    UseSystemPasswordChar = true
                };
                var okButton = new Button { Text = "确定", Location = new Point(176, 102), Size = new Size(76, 30) };
                var cancelButton = new Button { Text = "取消", Location = new Point(262, 102), Size = new Size(76, 30) };

                okButton.Click += (_, __) =>
                {
                    if (string.Equals(passwordBox.Text ?? string.Empty, LinkEditPassword, StringComparison.Ordinal))
                    {
                        dialog.DialogResult = DialogResult.OK;
                        dialog.Close();
                        return;
                    }

                    ShowErrorMessage("口令错误，未保存链接修改。", "设置");
                    passwordBox.Focus();
                    passwordBox.SelectAll();
                };
                cancelButton.Click += (_, __) =>
                {
                    dialog.DialogResult = DialogResult.Cancel;
                    dialog.Close();
                };

                dialog.AcceptButton = okButton;
                dialog.CancelButton = cancelButton;
                dialog.Controls.AddRange(new Control[] { label, passwordBox, okButton, cancelButton });
                dialog.Shown += (_, __) => passwordBox.Focus();
                return dialog.ShowDialog(this) == DialogResult.OK;
            }
        }

        private List<string> GetConfiguredOfficeTargetPaths()
        {
            var result = new List<string>();
            Action<string> addPath = path =>
            {
                var trimmed = (path ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(trimmed))
                {
                    return;
                }

                if (!result.Any(existing => string.Equals(existing, trimmed, StringComparison.OrdinalIgnoreCase)))
                {
                    result.Add(trimmed);
                }
            };

            foreach (var path in _settings.OfficeKeyTargetPaths)
            {
                addPath(path);
            }

            addPath(_settings.OfficeKeyTargetPath);

            if (result.Count == 0)
            {
                result.Add(GetDefaultOfficeTargetPath());
            }

            return result;
        }

        private static List<string> SplitMultiLinePaths(string text)
        {
            return (text ?? string.Empty)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Trim())
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private string ResolveOfficeTargetPath(string configuredPath)
        {
            var desktopDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (string.IsNullOrWhiteSpace(desktopDirectory))
            {
                desktopDirectory = GetApplicationDirectory();
            }

            var trimmed = (configuredPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return Path.Combine(desktopDirectory, OfficeKeyTargetPath);
            }

            var fileName = Path.GetFileName(trimmed);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                fileName = OfficeKeyTargetPath;
            }

            if (!Path.IsPathRooted(trimmed))
            {
                return Path.Combine(desktopDirectory, fileName);
            }

            if (IsLegacyDesktopTargetPath(trimmed, desktopDirectory))
            {
                return Path.Combine(desktopDirectory, fileName);
            }

            return trimmed;
        }

        private string SelectOfficeTargetPath()
        {
            var resolvedPaths = GetConfiguredOfficeTargetPaths()
                .Select(ResolveOfficeTargetPath)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var existingPath = resolvedPaths.FirstOrDefault(File.Exists);
            if (!string.IsNullOrWhiteSpace(existingPath))
            {
                return existingPath;
            }

            return resolvedPaths.FirstOrDefault() ?? GetDefaultOfficeTargetPath();
        }

        private void OpenOfficeKeySourceFile()
        {
            var sourcePath = BuildPortableCandidates(_settings.OfficeKeySourcePath, OfficeKeySourcePath).FirstOrDefault(File.Exists);
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                ShowErrorMessage("未找到 Office 密钥源文件。", "Office");
                return;
            }

            StartTarget(sourcePath);
        }

        private void OpenOfficeKeyTargetFile()
        {
            var targetPath = SelectOfficeTargetPath();
            if (string.IsNullOrWhiteSpace(targetPath) || !File.Exists(targetPath))
            {
                ShowErrorMessage("当前未找到桌面目标文件。", "Office");
                return;
            }

            StartTarget(targetPath);
        }

        private string GetOfficeSequenceMenuText()
        {
            var configuredText = (_settings.OfficePostActionCondition ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(configuredText)
                || string.Equals(configuredText, DefaultOfficePostActionCondition, StringComparison.OrdinalIgnoreCase))
            {
                return DefaultOfficeSequenceMenuText;
            }

            return configuredText;
        }

        private void RefreshOfficeSequenceMenuItem()
        {
            if (_officeSequenceMenuItem == null)
            {
                return;
            }

            var sequenceText = (_settings.OfficePostActionSequence ?? string.Empty).Trim();
            var showItem = _settings.OfficePostActionEnabled && !string.IsNullOrWhiteSpace(sequenceText);
            _officeSequenceMenuItem.Text = GetOfficeSequenceMenuText();
            _officeSequenceMenuItem.Visible = showItem;
            if (_officeSequenceMenuSeparator != null)
            {
                _officeSequenceMenuSeparator.Visible = showItem;
            }
        }

        private void ExecuteManualOfficeSequence()
        {
            var sequenceText = string.IsNullOrWhiteSpace(_settings.OfficePostActionSequence)
                ? DefaultOfficePostActionSequence
                : _settings.OfficePostActionSequence.Trim();
            if (string.IsNullOrWhiteSpace(sequenceText))
            {
                ShowErrorMessage("当前未设置 Office 右键按键串，请先到设置里的“Office按键”页填写。", "Office");
                return;
            }

            try
            {
                IntPtr officeWindowHandle;
                if (!WaitForOfficeMainWindow(out officeWindowHandle))
                {
                    string launchedTarget;
                    if (!TryLaunchOfficeApplication(out launchedTarget))
                    {
                        ShowErrorMessage("未检测到正在运行的 Office，也未能自动打开 Office 软件。", "Office");
                        return;
                    }

                    Thread.Sleep(1200);
                    if (!WaitForOfficeMainWindow(out officeWindowHandle))
                    {
                        ShowErrorMessage("Office 已启动，但暂时还没找到可接收按键的主窗口。", "Office");
                        return;
                    }
                }

                try
                {
                    ShowWindow(officeWindowHandle, 9);
                    SetForegroundWindow(officeWindowHandle);
                }
                catch
                {
                }

                Thread.Sleep(320);
                ExecuteOfficePostActionSequence(sequenceText);
            }
            catch (Exception ex)
            {
                ShowErrorMessage("发送 Office 按键串失败：" + ex.Message, "Office");
            }
        }

        private void ShowOfficeCommunicationTestDialog()
        {
            using (var dialog = new Form())
            {
                dialog.Text = "Office 通信测试";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.AutoScaleMode = AutoScaleMode.None;
                dialog.KeyPreview = true;
                dialog.ClientSize = new Size(700, 520);
                dialog.Font = Font;

                var titleLabel = new Label
                {
                    Text = "通信测试 + 人工辅助",
                    Location = new Point(18, 16),
                    Size = new Size(260, 24),
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point)
                };
                var noteBox = new TextBox
                {
                    Text = "本窗口只测试 URL 连通、发送模拟文本、接收模拟返回；全程不碰激活数据、不自动代填。\r\n按 Esc 可直接关闭窗口。",
                    Location = new Point(18, 44),
                    Size = new Size(664, 44),
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    Multiline = true,
                    BackColor = dialog.BackColor,
                    ForeColor = Color.DimGray,
                    TabStop = false
                };
                var urlLabel = new Label { Text = "测试 URL", Location = new Point(18, 96), Size = new Size(120, 22) };
                var urlBox = new TextBox
                {
                    Location = new Point(18, 122),
                    Size = new Size(664, 28),
                    Text = string.IsNullOrWhiteSpace(_settings.OfficeCommunicationTestUrl) ? DefaultOfficeCommunicationTestUrl : _settings.OfficeCommunicationTestUrl
                };
                var payloadLabel = new Label { Text = "模拟文本", Location = new Point(18, 160), Size = new Size(120, 22) };
                var payloadBox = new TextBox
                {
                    Location = new Point(18, 186),
                    Size = new Size(664, 68),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true,
                    Text = string.IsNullOrWhiteSpace(_settings.OfficeCommunicationTestPayload) ? DefaultOfficeCommunicationTestPayload : _settings.OfficeCommunicationTestPayload
                };
                var progressLabel = new Label { Text = "等待开始", Location = new Point(18, 266), Size = new Size(664, 22) };
                var progressBar = new ProgressBar
                {
                    Location = new Point(18, 292),
                    Size = new Size(664, 18),
                    Minimum = 0,
                    Maximum = 5,
                    Value = 0
                };
                var outputLabel = new Label { Text = "测试输出(可复制)", Location = new Point(18, 322), Size = new Size(140, 22) };
                var outputBox = new RichTextBox
                {
                    Location = new Point(18, 348),
                    Size = new Size(664, 122),
                    ReadOnly = true,
                    ShortcutsEnabled = true,
                    DetectUrls = false,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
                };
                var startButton = new Button { Text = "开始测试", Location = new Point(18, 480), Size = new Size(96, 30) };
                var copyButton = new Button { Text = "复制输出", Location = new Point(126, 480), Size = new Size(96, 30) };
                var closeButton = new Button { Text = "关闭", Location = new Point(586, 480), Size = new Size(96, 30) };

                Action<string> appendOutput = text =>
                {
                    if (string.IsNullOrWhiteSpace(text) || dialog.IsDisposed)
                    {
                        return;
                    }

                    try
                    {
                        dialog.BeginInvoke((MethodInvoker)delegate
                        {
                            if (dialog.IsDisposed || outputBox.IsDisposed)
                            {
                                return;
                            }

                            if (outputBox.TextLength > 0)
                            {
                                outputBox.AppendText(Environment.NewLine);
                            }

                            outputBox.AppendText(text);
                            outputBox.SelectionStart = outputBox.TextLength;
                            outputBox.ScrollToCaret();
                        });
                    }
                    catch
                    {
                    }
                };
                Action<int, string> setProgress = (step, statusText) =>
                {
                    if (dialog.IsDisposed)
                    {
                        return;
                    }

                    try
                    {
                        dialog.BeginInvoke((MethodInvoker)delegate
                        {
                            if (dialog.IsDisposed)
                            {
                                return;
                            }

                            progressBar.Value = Math.Max(progressBar.Minimum, Math.Min(progressBar.Maximum, step));
                            progressLabel.Text = statusText ?? string.Empty;
                        });
                    }
                    catch
                    {
                    }
                };
                Action<bool, string> finishTest = (success, summary) =>
                {
                    if (dialog.IsDisposed)
                    {
                        return;
                    }

                    try
                    {
                        dialog.BeginInvoke((MethodInvoker)delegate
                        {
                            if (dialog.IsDisposed)
                            {
                                return;
                            }

                            progressBar.Value = success ? progressBar.Maximum : Math.Min(progressBar.Maximum, Math.Max(progressBar.Value, 1));
                            progressLabel.Text = summary ?? string.Empty;
                            startButton.Enabled = true;
                        });
                    }
                    catch
                    {
                    }
                };

                Action startTest = delegate
                {
                    startButton.Enabled = false;
                    progressBar.Value = 0;
                    progressLabel.Text = "准备测试";
                    outputBox.Clear();

                    var urlText = (urlBox.Text ?? string.Empty).Trim();
                    var payloadText = payloadBox.Text ?? string.Empty;
                    _settings.OfficeCommunicationTestUrl = urlText;
                    _settings.OfficeCommunicationTestPayload = payloadText;
                    _settings.Save();

                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        Uri uri;
                        if (!Uri.TryCreate(urlText, UriKind.Absolute, out uri))
                        {
                            appendOutput("URL 无效：" + urlText);
                            finishTest(false, "URL 无效");
                            return;
                        }

                        setProgress(1, "1/5 正在检查 URL");
                        appendOutput("测试 URL: " + uri);

                        string probeSummary;
                        if (!TryProbeOfficeCommunicationUrl(uri, out probeSummary))
                        {
                            appendOutput("连通失败: " + probeSummary);
                            finishTest(false, "URL 连通失败");
                            return;
                        }

                        appendOutput("连通成功: " + probeSummary);
                        setProgress(2, "2/5 正在发送模拟文本");
                        appendOutput("发送模拟文本:");
                        appendOutput(payloadText);

                        string responseText;
                        string responseSummary;
                        var postSucceeded = TryPostOfficeCommunicationPayload(uri, payloadText, out responseText, out responseSummary);
                        setProgress(3, "3/5 正在接收模拟返回");
                        appendOutput("返回摘要: " + responseSummary);

                        if (!string.IsNullOrWhiteSpace(responseText))
                        {
                            appendOutput("返回内容:");
                            appendOutput(responseText);
                        }

                        if (!postSucceeded)
                        {
                            finishTest(false, "模拟文本发送失败");
                            return;
                        }

                        setProgress(4, "4/5 已接收模拟返回");
                        Thread.Sleep(180);
                        setProgress(5, "5/5 通信测试完成");
                        finishTest(true, "通信测试完成");
                    });
                };

                startButton.Click += (_, __) => startTest();
                copyButton.Click += (_, __) =>
                {
                    if (string.IsNullOrWhiteSpace(outputBox.Text))
                    {
                        return;
                    }

                    try
                    {
                        CopyTextToClipboard(outputBox.Text);
                    }
                    catch (Exception ex)
                    {
                        ShowErrorMessage("复制测试输出失败：" + ex.Message, "通信测试");
                    }
                };
                closeButton.Click += (_, __) => dialog.Close();
                dialog.KeyDown += (_, e) =>
                {
                    if (e != null && e.KeyCode == Keys.Escape)
                    {
                        e.SuppressKeyPress = true;
                        e.Handled = true;
                        dialog.Close();
                    }
                };
                dialog.Shown += (_, __) => startTest();

                dialog.Controls.AddRange(new Control[]
                {
                    titleLabel, noteBox, urlLabel, urlBox, payloadLabel, payloadBox,
                    progressLabel, progressBar, outputLabel, outputBox, startButton, copyButton, closeButton
                });

                dialog.ShowDialog(this);
            }
        }

        private bool TryProbeOfficeCommunicationUrl(Uri uri, out string summary)
        {
            summary = string.Empty;
            foreach (var method in new[] { "HEAD", "GET" })
            {
                HttpWebResponse response = null;
                try
                {
                    var request = (HttpWebRequest)WebRequest.Create(uri);
                    request.Method = method;
                    request.Timeout = 5000;
                    request.ReadWriteTimeout = 5000;
                    request.Proxy = null;
                    response = (HttpWebResponse)request.GetResponse();
                    summary = string.Format("{0} {1} ({2})", (int)response.StatusCode, response.StatusDescription, method);
                    return true;
                }
                catch (WebException ex)
                {
                    response = ex.Response as HttpWebResponse;
                    if (response != null)
                    {
                        summary = string.Format("{0} {1} ({2})", (int)response.StatusCode, response.StatusDescription, method);
                        return true;
                    }

                    summary = ex.Message;
                }
                catch (Exception ex)
                {
                    summary = ex.Message;
                }
                finally
                {
                    if (response != null)
                    {
                        response.Close();
                    }
                }
            }

            return false;
        }

        private bool TryPostOfficeCommunicationPayload(Uri uri, string payloadText, out string responseText, out string responseSummary)
        {
            responseText = string.Empty;
            responseSummary = string.Empty;

            if (TrySendOfficeCommunicationRequest(
                uri,
                "POST",
                "application/x-www-form-urlencoded; charset=utf-8",
                Encoding.UTF8.GetBytes("text=" + Uri.EscapeDataString(payloadText ?? string.Empty) + "&message=" + Uri.EscapeDataString(payloadText ?? string.Empty) + "&payload=" + Uri.EscapeDataString(payloadText ?? string.Empty)),
                "POST(form)",
                out responseText,
                out responseSummary))
            {
                return true;
            }

            if (TrySendOfficeCommunicationRequest(
                uri,
                "POST",
                "text/plain; charset=utf-8",
                Encoding.UTF8.GetBytes(payloadText ?? string.Empty),
                "POST(text)",
                out responseText,
                out responseSummary))
            {
                return true;
            }

            var queryUri = BuildOfficeCommunicationQueryUri(uri, payloadText ?? string.Empty);
            if (queryUri != null && TrySendOfficeCommunicationRequest(
                queryUri,
                "GET",
                null,
                null,
                "GET(query)",
                out responseText,
                out responseSummary))
            {
                return true;
            }

            return false;
        }

        private bool TrySendOfficeCommunicationRequest(Uri uri, string method, string contentType, byte[] bodyBytes, string modeTag, out string responseText, out string responseSummary)
        {
            responseText = string.Empty;
            responseSummary = string.Empty;
            HttpWebResponse response = null;
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(uri);
                request.Method = method;
                request.Timeout = 8000;
                request.ReadWriteTimeout = 8000;
                request.Proxy = null;
                request.UserAgent = "PcCheckTool/OfficeCommTest";
                if (!string.IsNullOrWhiteSpace(contentType))
                {
                    request.ContentType = contentType;
                }

                if (bodyBytes != null && bodyBytes.Length > 0)
                {
                    request.ContentLength = bodyBytes.Length;
                    using (var requestStream = request.GetRequestStream())
                    {
                        requestStream.Write(bodyBytes, 0, bodyBytes.Length);
                    }
                }

                response = (HttpWebResponse)request.GetResponse();
                responseSummary = string.Format("{0} {1} ({2})", (int)response.StatusCode, response.StatusDescription, modeTag);
                responseText = ReadOfficeCommunicationResponseText(response);
                return true;
            }
            catch (WebException ex)
            {
                response = ex.Response as HttpWebResponse;
                if (response != null)
                {
                    responseSummary = string.Format("{0} {1} ({2})", (int)response.StatusCode, response.StatusDescription, modeTag);
                    responseText = ReadOfficeCommunicationResponseText(response);
                }
                else
                {
                    responseSummary = string.Format("{0} ({1})", ex.Message, modeTag);
                }

                return false;
            }
            catch (Exception ex)
            {
                responseSummary = string.Format("{0} ({1})", ex.Message, modeTag);
                return false;
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        private static Uri BuildOfficeCommunicationQueryUri(Uri baseUri, string payloadText)
        {
            if (baseUri == null)
            {
                return null;
            }

            try
            {
                var builder = new UriBuilder(baseUri);
                var encodedPayload = Uri.EscapeDataString(payloadText ?? string.Empty);
                var existingQuery = (builder.Query ?? string.Empty).TrimStart('?');
                var addedQuery = "text=" + encodedPayload + "&message=" + encodedPayload + "&payload=" + encodedPayload;
                builder.Query = string.IsNullOrWhiteSpace(existingQuery) ? addedQuery : existingQuery + "&" + addedQuery;
                return builder.Uri;
            }
            catch
            {
                return null;
            }
        }

        private static string ReadOfficeCommunicationResponseText(WebResponse response)
        {
            if (response == null)
            {
                return string.Empty;
            }

            try
            {
                using (var stream = response.GetResponseStream())
                {
                    if (stream == null)
                    {
                        return string.Empty;
                    }

                    using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool IsLegacyDesktopTargetPath(string candidatePath, string currentDesktopDirectory)
        {
            if (string.IsNullOrWhiteSpace(candidatePath))
            {
                return false;
            }

            try
            {
                var fullCandidate = Path.GetFullPath(candidatePath);
                var currentDesktop = string.IsNullOrWhiteSpace(currentDesktopDirectory)
                    ? string.Empty
                    : Path.GetFullPath(currentDesktopDirectory);

                if (!string.IsNullOrWhiteSpace(currentDesktop)
                    && string.Equals(fullCandidate, currentDesktop, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                var directory = Path.GetDirectoryName(fullCandidate) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(currentDesktop)
                    && string.Equals(directory, currentDesktop, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                var upper = fullCandidate.ToUpperInvariant();
                return upper.Contains(@"\DESKTOP\")
                    || upper.Contains(@"\桌面\")
                    || upper.Contains(@"\デスクトップ\")
                    || upper.Contains(@"\ONEDRIVE\DESKTOP\");
            }
            catch
            {
                return false;
            }
        }

        private List<string> BuildOfficeApplicationCandidates()
        {
            var result = new List<string>();
            Action<string> addCandidate = candidate =>
            {
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    return;
                }

                if (!result.Any(existing => string.Equals(existing, candidate, StringComparison.OrdinalIgnoreCase)))
                {
                    result.Add(candidate);
                }
            };

            foreach (var configuredPath in _settings.OfficeAppPaths
                .Concat(new[] { _settings.OfficeAppPath })
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                addCandidate(ResolveConfiguredPath(configuredPath));
                addCandidate(configuredPath);
            }

            foreach (var exeName in new[] { "WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE" })
            {
                foreach (var root in new[] { Registry.CurrentUser, Registry.LocalMachine })
                {
                    foreach (var subKeyPath in new[]
                    {
                        string.Format(@"Software\Microsoft\Windows\CurrentVersion\App Paths\{0}", exeName),
                        string.Format(@"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\{0}", exeName)
                    })
                    {
                        using (var subKey = root.OpenSubKey(subKeyPath))
                        {
                            var registryPath = subKey == null ? null : subKey.GetValue(string.Empty) as string;
                            addCandidate(registryPath);
                        }
                    }
                }
            }

            foreach (var programRoot in new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            }.Where(path => !string.IsNullOrWhiteSpace(path)))
            {
                foreach (var officeFolder in new[]
                {
                    @"Microsoft Office\root\Office16",
                    @"Microsoft Office\Office16",
                    @"Microsoft Office\root\Office15",
                    @"Microsoft Office\Office15",
                    @"Microsoft Office\root\Office14",
                    @"Microsoft Office\Office14"
                })
                {
                    foreach (var exeName in new[] { "WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE" })
                    {
                        addCandidate(Path.Combine(programRoot, officeFolder, exeName));
                    }
                }
            }

            addCandidate("WINWORD.EXE");
            addCandidate("EXCEL.EXE");
            addCandidate("POWERPNT.EXE");
            return result;
        }

        private bool TryLaunchSpecificOfficeApplication(string executableName, out string launchedTarget)
        {
            launchedTarget = null;
            var normalizedExe = (executableName ?? string.Empty).Trim();
            foreach (var candidate in BuildOfficeApplicationCandidates())
            {
                var candidateFileName = Path.GetFileName(candidate) ?? candidate;
                if (!string.IsNullOrWhiteSpace(normalizedExe)
                    && !string.Equals(candidateFileName, normalizedExe, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (Path.IsPathRooted(candidate) && !File.Exists(candidate))
                {
                    continue;
                }

                if (TryStartTarget(candidate))
                {
                    launchedTarget = candidate;
                    return true;
                }
            }

            return false;
        }

        private bool TryLaunchOfficeApplication(out string launchedTarget)
        {
            foreach (var candidate in BuildOfficeApplicationCandidates())
            {
                if (Path.IsPathRooted(candidate) && !File.Exists(candidate))
                {
                    continue;
                }

                if (TryStartTarget(candidate))
                {
                    launchedTarget = candidate;
                    return true;
                }
            }

            launchedTarget = null;
            return false;
        }

        private void BeginOfficeActivationUiAutomation(string officeKey)
        {
            if (!_settings.OfficePostActionEnabled)
            {
                return;
            }

            var keyToPaste = (officeKey ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(keyToPaste))
            {
                return;
            }

            var worker = new Thread(() =>
            {
                try
                {
                    Thread.Sleep(1800);
                    IntPtr windowHandle;
                    if (!WaitForOfficeMainWindow(out windowHandle))
                    {
                        return;
                    }

                    try
                    {
                        ShowWindow(windowHandle, 9);
                        SetForegroundWindow(windowHandle);
                    }
                    catch
                    {
                    }

                    Thread.Sleep(420);
                    var matchedRule = WaitForOfficePostActionRule(windowHandle, _settings.OfficePostActionCondition, _settings.OfficePostActionSequence);
                    if (matchedRule == null)
                    {
                        return;
                    }

                    try
                    {
                        ShowWindow(windowHandle, 9);
                        SetForegroundWindow(windowHandle);
                    }
                    catch
                    {
                    }

                    ExecuteOfficePostActionSequence(matchedRule.Sequence);
                    TryExecuteOfficePostAhkScript(keyToPaste, matchedRule.Condition);
                }
                catch
                {
                }
            });
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }

        private bool TryExecuteOfficePostAhkScript(string officeKey, string matchedCondition)
        {
            if (!_settings.OfficePostAhkEnabled)
            {
                return false;
            }

            var scriptBody = (_settings.OfficePostAhkScript ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(scriptBody))
            {
                return false;
            }

            var ahkPath = BuildPortableCandidates(_settings.OfficePostAhkPath, "AutoHotkey.exe", "AutoHotkey64.exe")
                .FirstOrDefault(File.Exists);
            if (string.IsNullOrWhiteSpace(ahkPath) || !File.Exists(ahkPath))
            {
                return false;
            }

            var expandedScript = scriptBody
                .Replace("{key}", officeKey ?? string.Empty)
                .Replace("{condition}", matchedCondition ?? string.Empty)
                .Replace("{appdir}", GetApplicationDirectory());
            var tempScriptPath = Path.Combine(Path.GetTempPath(), "PcCheckTool_OfficePost_" + Guid.NewGuid().ToString("N") + ".ahk");
            File.WriteAllText(tempScriptPath, expandedScript, new UTF8Encoding(true));

            var startInfo = new ProcessStartInfo
            {
                FileName = ahkPath,
                Arguments = "\"" + tempScriptPath + "\"",
                UseShellExecute = true,
                WorkingDirectory = Path.GetDirectoryName(ahkPath) ?? GetApplicationDirectory()
            };

            try
            {
                Process.Start(startInfo);
            }
            catch
            {
                return false;
            }

            ThreadPool.QueueUserWorkItem(_ =>
            {
                Thread.Sleep(30000);
                try
                {
                    if (File.Exists(tempScriptPath))
                    {
                        File.Delete(tempScriptPath);
                    }
                }
                catch
                {
                }
            });

            return true;
        }

        private OfficePostActionRule WaitForOfficePostActionRule(IntPtr mainWindowHandle, string conditionText, string sequenceText)
        {
            var triggerStates = BuildOfficePostTriggerStates(conditionText, sequenceText).ToList();
            if (triggerStates.Count == 0)
            {
                return null;
            }

            var needsPopupMonitoring = triggerStates.Any(state => state.Kind == OfficePostTriggerKind.PopupClosed);
            for (var i = 0; i < 1200; i++)
            {
                if (IsOfficeAlreadyActivated())
                {
                    return null;
                }

                var popupClosedThisLoop = false;
                if (needsPopupMonitoring)
                {
                    popupClosedThisLoop = TryCloseOfficePopupWindowsOnce(mainWindowHandle);
                }

                foreach (var state in triggerStates)
                {
                    if (IsOfficeTriggerStateSatisfied(state, popupClosedThisLoop))
                    {
                        return state.Rule;
                    }
                }

                Thread.Sleep(120);
            }

            return null;
        }

        private IEnumerable<OfficePostTriggerState> BuildOfficePostTriggerStates(string conditionText, string sequenceText)
        {
            foreach (var rule in BuildOfficePostActionRules(conditionText, sequenceText))
            {
                yield return CreateOfficePostTriggerState(rule);
            }
        }

        private IEnumerable<OfficePostActionRule> BuildOfficePostActionRules(string conditionText, string sequenceText)
        {
            var parsedRules = ParseOfficePostActionRulesFromSequence(sequenceText).ToList();
            if (parsedRules.Count > 0)
            {
                return parsedRules;
            }

            var normalizedSequence = string.IsNullOrWhiteSpace(sequenceText)
                ? DefaultOfficePostActionSequence
                : sequenceText.Trim();
            var conditions = SplitOfficeTriggerConditions(conditionText).ToList();
            if (conditions.Count == 0)
            {
                conditions.Add(DefaultOfficePostActionCondition);
            }

            return conditions
                .Select(condition => new OfficePostActionRule
                {
                    Condition = condition,
                    Sequence = normalizedSequence
                })
                .ToArray();
        }

        private IEnumerable<OfficePostActionRule> ParseOfficePostActionRulesFromSequence(string sequenceText)
        {
            var lines = (sequenceText ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Split(new[] { '\n' }, StringSplitOptions.None);
            var hasRuleHeader = false;
            foreach (var line in lines)
            {
                string ignoredCondition;
                if (TryParseOfficeRuleConditionHeader(line, out ignoredCondition))
                {
                    hasRuleHeader = true;
                    break;
                }
            }
            if (!hasRuleHeader)
            {
                return Enumerable.Empty<OfficePostActionRule>();
            }

            var rules = new List<OfficePostActionRule>();
            string currentCondition = null;
            var currentSequenceLines = new List<string>();
            Action flushRule = () =>
            {
                if (string.IsNullOrWhiteSpace(currentCondition) || currentSequenceLines.Count == 0)
                {
                    currentSequenceLines.Clear();
                    return;
                }

                rules.Add(new OfficePostActionRule
                {
                    Condition = currentCondition,
                    Sequence = string.Join(Environment.NewLine, currentSequenceLines)
                });
                currentSequenceLines.Clear();
            };

            foreach (var rawLine in lines)
            {
                var line = (rawLine ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                string ruleCondition;
                if (TryParseOfficeRuleConditionHeader(line, out ruleCondition))
                {
                    flushRule();
                    currentCondition = ruleCondition;
                    continue;
                }

                currentSequenceLines.Add(line);
            }

            flushRule();
            return rules;
        }

        private static bool TryParseOfficeRuleConditionHeader(string line, out string condition)
        {
            condition = string.Empty;
            var text = (line ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var separatorIndex = text.IndexOfAny(new[] { ':', '：' });
            if (separatorIndex <= 0)
            {
                return false;
            }

            var head = text.Substring(0, separatorIndex).Trim();
            if (!string.Equals(head, "条件", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(head, "触发", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(head, "when", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            condition = text.Substring(separatorIndex + 1).Trim();
            return !string.IsNullOrWhiteSpace(condition);
        }

        private static IEnumerable<string> SplitOfficeTriggerConditions(string conditionText)
        {
            var text = (conditionText ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');
            if (string.IsNullOrWhiteSpace(text))
            {
                return Enumerable.Empty<string>();
            }

            var result = new List<string>();
            foreach (var rawLine in text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                foreach (var rawPart in rawLine.Split(new[] { '|', ';', '；' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var normalized = (rawPart ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(normalized))
                    {
                        result.Add(normalized);
                    }
                }
            }

            return result
                .Select(NormalizeOfficePostActionConditionText)
                .Where(condition => !string.IsNullOrWhiteSpace(condition))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static OfficePostTriggerState CreateOfficePostTriggerState(OfficePostActionRule rule)
        {
            var state = new OfficePostTriggerState
            {
                Rule = rule ?? new OfficePostActionRule(),
                StartedAt = Stopwatch.StartNew()
            };
            var normalizedCondition = NormalizeOfficePostActionConditionText(state.Rule.Condition);
            if (string.IsNullOrWhiteSpace(normalizedCondition))
            {
                normalizedCondition = DefaultOfficePostActionCondition;
            }

            Keys triggerKey;
            if (TryParseOfficePostTriggerKey(normalizedCondition, out triggerKey))
            {
                state.Kind = OfficePostTriggerKind.Key;
                state.TriggerKey = triggerKey;
                return state;
            }

            if (normalizedCondition.IndexOf("立即", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                state.Kind = OfficePostTriggerKind.Immediate;
                return state;
            }

            if (normalizedCondition.IndexOf("弹窗", StringComparison.OrdinalIgnoreCase) >= 0
                || normalizedCondition.IndexOf("关闭", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                state.Kind = OfficePostTriggerKind.PopupClosed;
                return state;
            }

            var delayMs = ExtractDelayFromCondition(normalizedCondition);
            if (delayMs > 0)
            {
                state.Kind = OfficePostTriggerKind.Delay;
                state.DelayMs = delayMs;
                return state;
            }

            state.Kind = OfficePostTriggerKind.Immediate;
            return state;
        }

        private bool IsOfficeTriggerStateSatisfied(OfficePostTriggerState state, bool popupClosedThisLoop)
        {
            if (state == null)
            {
                return false;
            }

            switch (state.Kind)
            {
                case OfficePostTriggerKind.Immediate:
                    return true;
                case OfficePostTriggerKind.Delay:
                    return state.StartedAt != null && state.StartedAt.ElapsedMilliseconds >= state.DelayMs;
                case OfficePostTriggerKind.PopupClosed:
                    return popupClosedThisLoop;
                case OfficePostTriggerKind.Key:
                    var isDown = (GetAsyncKeyState((int)state.TriggerKey) & 0x8000) != 0;
                    var triggered = isDown && !state.PreviousDown;
                    state.PreviousDown = isDown;
                    return triggered;
                default:
                    return false;
            }
        }

        private bool TryCloseOfficePopupWindowsOnce(IntPtr mainWindowHandle)
        {
            var popupWindows = GetVisibleOfficePopupWindows(mainWindowHandle).ToList();
            if (popupWindows.Count == 0)
            {
                return false;
            }

            foreach (var popupWindow in popupWindows)
            {
                try
                {
                    PostMessage(popupWindow, 0x0010, IntPtr.Zero, IntPtr.Zero);
                }
                catch
                {
                }
            }

            Thread.Sleep(320);
            return !GetVisibleOfficePopupWindows(mainWindowHandle).Any();
        }

        private bool IsOfficeAlreadyActivated()
        {
            bool activated;
            string statusText;
            return TryGetOfficeActivationState(out activated, out statusText) && activated;
        }

        private static string NormalizeOfficePostActionConditionText(string conditionText)
        {
            var text = (conditionText ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            return text
                .Replace(" ", string.Empty)
                .Replace("　", string.Empty)
                .Replace("按下", "按")
                .Replace("之后", "后")
                .Replace("以后", "后");
        }

        private static int ExtractDelayFromCondition(string conditionText)
        {
            var match = Regex.Match(conditionText ?? string.Empty, @"(\d+)");
            int delayMs;
            if (match.Success && int.TryParse(match.Groups[1].Value, out delayMs))
            {
                return Math.Max(0, delayMs);
            }

            return 0;
        }

        private static bool TryParseOfficePostTriggerKey(string conditionText, out Keys triggerKey)
        {
            triggerKey = Keys.None;
            var normalized = NormalizeOfficePostActionConditionText(conditionText);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            var match = Regex.Match(normalized, @"用户按(.+?)后?$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                match = Regex.Match(normalized, @"按(.+?)后?$", RegexOptions.IgnoreCase);
            }

            if (!match.Success)
            {
                return false;
            }

            var token = match.Groups[1].Value;
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            return TryParseOfficeTriggerKeyToken(token, out triggerKey);
        }

        private static bool TryParseOfficeTriggerKeyToken(string token, out Keys triggerKey)
        {
            triggerKey = Keys.None;
            var normalized = (token ?? string.Empty).Trim().Trim('{', '}');
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            var upper = normalized.ToUpperInvariant();
            if (upper.Length == 1 && upper[0] >= 'A' && upper[0] <= 'Z')
            {
                triggerKey = (Keys)Enum.Parse(typeof(Keys), upper, true);
                return true;
            }

            if (upper.Length == 1 && upper[0] >= '0' && upper[0] <= '9')
            {
                triggerKey = Keys.D0 + (upper[0] - '0');
                return true;
            }

            switch (upper)
            {
                case "ENTER":
                case "RETURN":
                case "回车":
                    triggerKey = Keys.Enter;
                    return true;
                case "ESC":
                case "ESCAPE":
                    triggerKey = Keys.Escape;
                    return true;
                case "TAB":
                    triggerKey = Keys.Tab;
                    return true;
                case "SPACE":
                case "空格":
                    triggerKey = Keys.Space;
                    return true;
                case "UP":
                    triggerKey = Keys.Up;
                    return true;
                case "DOWN":
                    triggerKey = Keys.Down;
                    return true;
                case "LEFT":
                    triggerKey = Keys.Left;
                    return true;
                case "RIGHT":
                    triggerKey = Keys.Right;
                    return true;
                case "DELETE":
                case "DEL":
                    triggerKey = Keys.Delete;
                    return true;
                case "BACKSPACE":
                    triggerKey = Keys.Back;
                    return true;
                case "HOME":
                    triggerKey = Keys.Home;
                    return true;
                case "END":
                    triggerKey = Keys.End;
                    return true;
                case "PAGEUP":
                case "PGUP":
                    triggerKey = Keys.PageUp;
                    return true;
                case "PAGEDOWN":
                case "PGDN":
                    triggerKey = Keys.PageDown;
                    return true;
                default:
                    if (Regex.IsMatch(upper, @"^F([1-9]|1[0-2])$"))
                    {
                        triggerKey = (Keys)Enum.Parse(typeof(Keys), upper, true);
                        return true;
                    }

                    return false;
            }
        }

        private void WaitAndCloseOfficePopupWindows(IntPtr mainWindowHandle)
        {
            for (var attempt = 0; attempt < 12; attempt++)
            {
                var popupWindows = GetVisibleOfficePopupWindows(mainWindowHandle).ToList();
                if (popupWindows.Count == 0)
                {
                    Thread.Sleep(220);
                    popupWindows = GetVisibleOfficePopupWindows(mainWindowHandle).ToList();
                    if (popupWindows.Count == 0)
                    {
                        return;
                    }
                }

                foreach (var popupWindow in popupWindows)
                {
                    try
                    {
                        PostMessage(popupWindow, 0x0010, IntPtr.Zero, IntPtr.Zero);
                    }
                    catch
                    {
                    }
                }

                Thread.Sleep(450);
            }
        }

        private IEnumerable<IntPtr> GetVisibleOfficePopupWindows(IntPtr mainWindowHandle)
        {
            var processIds = GetOfficeProcessIds();
            if (processIds.Count == 0)
            {
                return Enumerable.Empty<IntPtr>();
            }

            var windows = new List<IntPtr>();
            EnumWindows((hWnd, _) =>
            {
                try
                {
                    if (hWnd == IntPtr.Zero || hWnd == mainWindowHandle || !IsWindowVisible(hWnd))
                    {
                        return true;
                    }

                    uint processId;
                    GetWindowThreadProcessId(hWnd, out processId);
                    if (!processIds.Contains((int)processId))
                    {
                        return true;
                    }

                    windows.Add(hWnd);
                }
                catch
                {
                }

                return true;
            }, IntPtr.Zero);

            return windows;
        }

        private static HashSet<int> GetOfficeProcessIds()
        {
            var result = new HashSet<int>();
            foreach (var processName in new[] { "WINWORD", "EXCEL", "POWERPNT" })
            {
                Process[] processes;
                try
                {
                    processes = Process.GetProcessesByName(processName);
                }
                catch
                {
                    continue;
                }

                foreach (var process in processes)
                {
                    try
                    {
                        if (!process.HasExited)
                        {
                            result.Add(process.Id);
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }

            return result;
        }

        private void ExecuteOfficePostActionSequence(string sequence)
        {
            foreach (var step in ParseOfficePostActionSequence(sequence))
            {
                if (step.DelayMs > 0)
                {
                    Thread.Sleep(step.DelayMs);
                }

                if (step.VirtualKeyCode != 0)
                {
                    TapVirtualKey(step.VirtualKeyCode);
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(step.Keys))
                {
                    SendKeys.SendWait(step.Keys);
                }
            }
        }

        private IEnumerable<OfficePostActionStep> ParseOfficePostActionSequence(string sequence)
        {
            var rawTokens = SplitOfficeSequenceParts(sequence).ToList();

            if (rawTokens.Count == 0)
            {
                rawTokens.AddRange(SplitOfficeSequenceParts(DefaultOfficePostActionSequence));
            }

            var steps = new List<OfficePostActionStep>();
            for (var index = 0; index < rawTokens.Count; index++)
            {
                var token = rawTokens[index];
                var lowerToken = token.ToLowerInvariant();

                if (lowerToken.StartsWith("按键:") || lowerToken.StartsWith("key:"))
                {
                    var keyBody = token.Substring(token.IndexOf(':') + 1).Trim();
                    foreach (var keyToken in SplitOfficeSequenceParts(keyBody))
                    {
                        var parsedKeyToken = ParseOfficePostActionToken(keyToken);
                        if (parsedKeyToken.DelayMs > 0 || parsedKeyToken.VirtualKeyCode != 0 || !string.IsNullOrWhiteSpace(parsedKeyToken.Keys))
                        {
                            steps.Add(parsedKeyToken);
                        }
                    }

                    continue;
                }

                if (lowerToken.StartsWith("停顿:") || lowerToken.StartsWith("sleep:") || lowerToken.StartsWith("延时:"))
                {
                    var delayBody = token.Substring(token.IndexOf(':') + 1).Trim();
                    int delayMs;
                    if (int.TryParse(Regex.Match(delayBody, @"(\d+)").Value, out delayMs) && delayMs > 0)
                    {
                        steps.Add(new OfficePostActionStep { DelayMs = delayMs });
                    }

                    continue;
                }

                if (lowerToken.StartsWith("文本:") || lowerToken.StartsWith("text:"))
                {
                    var textBody = token.Substring(token.IndexOf(':') + 1);
                    steps.Add(new OfficePostActionStep { Keys = EscapeOfficeSendKeysLiteralText(textBody) });
                    continue;
                }

                if (lowerToken.StartsWith("条件:") || lowerToken.StartsWith("触发:") || lowerToken.StartsWith("when:"))
                {
                    continue;
                }

                var parsedToken = ParseOfficePostActionToken(token);
                if (parsedToken.DelayMs > 0 || parsedToken.VirtualKeyCode != 0 || !string.IsNullOrWhiteSpace(parsedToken.Keys))
                {
                    steps.Add(parsedToken);
                }
            }

            return steps.Where(step => step.DelayMs > 0 || step.VirtualKeyCode != 0 || !string.IsNullOrWhiteSpace(step.Keys)).ToArray();
        }

        private OfficePostActionStep ParseOfficePostActionToken(string token)
        {
            var lowerToken = (token ?? string.Empty).Trim().ToLowerInvariant();
            if (lowerToken.StartsWith("sleep") || lowerToken.StartsWith("delay") || lowerToken.StartsWith("延时") || lowerToken.StartsWith("等待"))
            {
                var match = Regex.Match(lowerToken, @"(\d+)");
                int delayMs;
                if (match.Success && int.TryParse(match.Groups[1].Value, out delayMs))
                {
                    return new OfficePostActionStep { DelayMs = Math.Max(0, delayMs) };
                }

                return new OfficePostActionStep();
            }

            if (string.Equals(lowerToken, "alt", StringComparison.Ordinal))
            {
                return new OfficePostActionStep { VirtualKeyCode = 0x12 };
            }

            if (string.Equals(lowerToken, "ctrl", StringComparison.Ordinal) || string.Equals(lowerToken, "control", StringComparison.Ordinal))
            {
                return new OfficePostActionStep { VirtualKeyCode = 0x11 };
            }

            if (string.Equals(lowerToken, "shift", StringComparison.Ordinal))
            {
                return new OfficePostActionStep { VirtualKeyCode = 0x10 };
            }

            byte standaloneVirtualKey;
            if (TryParseStandaloneAhkVirtualKeyToken(token, out standaloneVirtualKey))
            {
                return new OfficePostActionStep { VirtualKeyCode = standaloneVirtualKey };
            }

            return new OfficePostActionStep { Keys = ConvertOfficePostActionToken(token, string.Empty) };
        }

        private static IEnumerable<string> SplitOfficeSequenceParts(string sequence)
        {
            var text = sequence ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                return Enumerable.Empty<string>();
            }

            if (text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0)
            {
                return text
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(token => (token ?? string.Empty).Trim())
                    .Where(token => !string.IsNullOrWhiteSpace(token))
                    .ToArray();
            }

            return text
                .Split(new[] { "--" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(token => (token ?? string.Empty).Trim())
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .ToArray();
        }

        private static string ConvertOfficePostActionToken(string token, string prefix)
        {
            var normalized = (token ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            var lower = normalized.ToLowerInvariant();
            if (lower.StartsWith("文本:"))
            {
                return EscapeOfficeSendKeysLiteralText(normalized.Substring(3));
            }

            if (lower.StartsWith("text:"))
            {
                return EscapeOfficeSendKeysLiteralText(normalized.Substring(5));
            }

            if (lower.StartsWith("发送文本:"))
            {
                return EscapeOfficeSendKeysLiteralText(normalized.Substring(5));
            }

            if (lower.StartsWith("{text}"))
            {
                return EscapeOfficeSendKeysLiteralText(normalized.Substring(6));
            }

            if (lower.StartsWith("tab*"))
            {
                int count;
                if (int.TryParse(lower.Substring(4), out count) && count > 0)
                {
                    return string.Concat(Enumerable.Repeat("{TAB}", count));
                }
            }

            string repeatedSpecialKey;
            if (TryConvertAhkRepeatedKey(normalized, out repeatedSpecialKey))
            {
                return repeatedSpecialKey;
            }

            if (lower.StartsWith("ctrl+"))
            {
                return "^" + ConvertSingleOfficeSendKey(lower.Substring(5));
            }

            if (lower.StartsWith("alt+"))
            {
                return "%" + ConvertSingleOfficeSendKey(normalized.Substring(4));
            }

            if (lower.StartsWith("shift+"))
            {
                return "+" + ConvertSingleOfficeSendKey(lower.Substring(6));
            }

            string ahkStyleKeys;
            if (TryConvertAhkModifierToken(normalized, out ahkStyleKeys))
            {
                return ahkStyleKeys;
            }

            return prefix + ConvertSingleOfficeSendKey(normalized);
        }

        private static string ConvertSingleOfficeSendKey(string token)
        {
            var lower = (token ?? string.Empty).Trim().ToLowerInvariant();
            switch (lower)
            {
                case "tab":
                    return "{TAB}";
                case "回车":
                case "enter":
                    return "{ENTER}";
                case "esc":
                case "escape":
                    return "{ESC}";
                case "up":
                    return "{UP}";
                case "down":
                    return "{DOWN}";
                case "left":
                    return "{LEFT}";
                case "right":
                    return "{RIGHT}";
                case "delete":
                case "del":
                    return "{DEL}";
                case "backspace":
                    return "{BACKSPACE}";
                case "home":
                    return "{HOME}";
                case "end":
                    return "{END}";
                case "pageup":
                case "pgup":
                    return "{PGUP}";
                case "pagedown":
                case "pgdn":
                    return "{PGDN}";
                case "space":
                case "空格":
                    return " ";
                default:
                    if (Regex.IsMatch(lower, @"^f([1-9]|1[0-2])$"))
                    {
                        return "{" + lower.ToUpperInvariant() + "}";
                    }

                    return token;
            }
        }

        private static string BuildOfficePostTokenFromKeyEvent(KeyEventArgs e)
        {
            if (e == null)
            {
                return string.Empty;
            }

            var keyCode = e.KeyCode;
            if (keyCode == Keys.None || keyCode == Keys.ControlKey || keyCode == Keys.ShiftKey || keyCode == Keys.Menu)
            {
                return string.Empty;
            }

            var token = NormalizeOfficePostKeyToken(keyCode);
            if (string.IsNullOrWhiteSpace(token))
            {
                return string.Empty;
            }

            if (e.Alt)
            {
                return "{Alt}--" + token;
            }

            if (e.Control)
            {
                return "^" + token;
            }

            if (e.Shift)
            {
                return "+" + token;
            }

            return token;
        }

        private static bool TryBuildStandaloneOfficeModifierToken(KeyEventArgs e, out string token)
        {
            token = string.Empty;
            if (e == null)
            {
                return false;
            }

            switch (e.KeyCode)
            {
                case Keys.Menu:
                    token = "{Alt}";
                    return true;
                case Keys.ControlKey:
                    token = "{Ctrl}";
                    return true;
                case Keys.ShiftKey:
                    token = "{Shift}";
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryParseStandaloneAhkVirtualKeyToken(string token, out byte virtualKeyCode)
        {
            virtualKeyCode = 0;
            var normalized = (token ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            switch (normalized.ToLowerInvariant())
            {
                case "{alt}":
                    virtualKeyCode = 0x12;
                    return true;
                case "{ctrl}":
                case "{control}":
                    virtualKeyCode = 0x11;
                    return true;
                case "{shift}":
                    virtualKeyCode = 0x10;
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryConvertAhkModifierToken(string token, out string sendKeysToken)
        {
            sendKeysToken = string.Empty;
            var normalized = (token ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            var modifierPrefix = new StringBuilder();
            var index = 0;
            while (index < normalized.Length)
            {
                switch (normalized[index])
                {
                    case '!':
                        modifierPrefix.Append('%');
                        index++;
                        continue;
                    case '^':
                        modifierPrefix.Append('^');
                        index++;
                        continue;
                    case '+':
                        modifierPrefix.Append('+');
                        index++;
                        continue;
                    default:
                        break;
                }
                break;
            }

            if (modifierPrefix.Length == 0 || index >= normalized.Length)
            {
                return false;
            }

            var remainder = normalized.Substring(index);
            if (string.IsNullOrWhiteSpace(remainder))
            {
                return false;
            }

            if (remainder.StartsWith("{", StringComparison.Ordinal) && remainder.EndsWith("}", StringComparison.Ordinal))
            {
                var innerToken = remainder.Substring(1, remainder.Length - 2);
                sendKeysToken = modifierPrefix + ConvertSingleOfficeSendKey(innerToken);
                return true;
            }

            sendKeysToken = modifierPrefix + ConvertSingleOfficeSendKey(remainder);
            return true;
        }

        private static bool TryConvertAhkRepeatedKey(string token, out string sendKeysToken)
        {
            sendKeysToken = string.Empty;
            var normalized = (token ?? string.Empty).Trim();
            var match = Regex.Match(normalized, @"^\{([A-Za-z][A-Za-z0-9]*)\s+(\d+)\}$");
            if (!match.Success)
            {
                return false;
            }

            int count;
            if (!int.TryParse(match.Groups[2].Value, out count) || count <= 0)
            {
                return false;
            }

            var baseKey = ConvertSingleOfficeSendKey(match.Groups[1].Value);
            if (string.IsNullOrWhiteSpace(baseKey))
            {
                return false;
            }

            sendKeysToken = string.Concat(Enumerable.Repeat(baseKey, count));
            return true;
        }

        private static string NormalizeOfficePostKeyToken(Keys keyCode)
        {
            if (keyCode >= Keys.A && keyCode <= Keys.Z)
            {
                return keyCode.ToString().ToUpperInvariant();
            }

            if (keyCode >= Keys.D0 && keyCode <= Keys.D9)
            {
                return ((char)('0' + (keyCode - Keys.D0))).ToString();
            }

            if (keyCode >= Keys.NumPad0 && keyCode <= Keys.NumPad9)
            {
                return ((char)('0' + (keyCode - Keys.NumPad0))).ToString();
            }

            switch (keyCode)
            {
                case Keys.Tab:
                    return "{Tab}";
                case Keys.Enter:
                    return "{Enter}";
                case Keys.Escape:
                    return "{Esc}";
                case Keys.Space:
                    return "{Space}";
                case Keys.Up:
                    return "{Up}";
                case Keys.Down:
                    return "{Down}";
                case Keys.Left:
                    return "{Left}";
                case Keys.Right:
                    return "{Right}";
                case Keys.Delete:
                    return "{Delete}";
                case Keys.Back:
                    return "{Backspace}";
                case Keys.Home:
                    return "{Home}";
                case Keys.End:
                    return "{End}";
                case Keys.PageUp:
                    return "{PgUp}";
                case Keys.PageDown:
                    return "{PgDn}";
                case Keys.OemMinus:
                    return "-";
                case Keys.Oemplus:
                    return "=";
                case Keys.Oemcomma:
                    return ",";
                case Keys.OemPeriod:
                    return ".";
                case Keys.OemQuestion:
                    return "/";
                case Keys.OemSemicolon:
                    return ";";
                case Keys.OemQuotes:
                    return "'";
                case Keys.OemOpenBrackets:
                    return "[";
                case Keys.OemCloseBrackets:
                    return "]";
                case Keys.OemPipe:
                    return "\\";
                case Keys.Oemtilde:
                    return "`";
                default:
                    if (keyCode >= Keys.F1 && keyCode <= Keys.F12)
                    {
                        return "{" + keyCode.ToString().ToUpperInvariant() + "}";
                    }

                    return keyCode.ToString();
            }
        }

        private static string EscapeOfficeSendKeysLiteralText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(text.Length * 2);
            foreach (var ch in text)
            {
                switch (ch)
                {
                    case '{':
                        builder.Append("{{}");
                        break;
                    case '}':
                        builder.Append("{}}");
                        break;
                    case '+':
                    case '^':
                    case '%':
                    case '~':
                    case '(':
                    case ')':
                    case '[':
                    case ']':
                        builder.Append('{').Append(ch).Append('}');
                        break;
                    case '\t':
                        builder.Append("{TAB}");
                        break;
                    case '\r':
                        break;
                    case '\n':
                        builder.Append("{ENTER}");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }

            return builder.ToString();
        }

        private static void TapVirtualKey(byte virtualKeyCode)
        {
            try
            {
                keybd_event(virtualKeyCode, 0, 0, UIntPtr.Zero);
                Thread.Sleep(40);
                keybd_event(virtualKeyCode, 0, KeyeventfKeyup, UIntPtr.Zero);
            }
            catch
            {
            }
        }

        private static bool WaitForOfficeMainWindow(out IntPtr windowHandle)
        {
            var processNames = new[] { "WINWORD", "EXCEL", "POWERPNT" };
            for (var attempt = 0; attempt < 30; attempt++)
            {
                foreach (var processName in processNames)
                {
                    Process[] processes;
                    try
                    {
                        processes = Process.GetProcessesByName(processName);
                    }
                    catch
                    {
                        continue;
                    }

                    foreach (var process in processes.OrderByDescending(item =>
                    {
                        try
                        {
                            return item.StartTime;
                        }
                        catch
                        {
                            return DateTime.MinValue;
                        }
                    }))
                    {
                        try
                        {
                            if (!process.HasExited && process.MainWindowHandle != IntPtr.Zero)
                            {
                                windowHandle = process.MainWindowHandle;
                                return true;
                            }
                        }
                        catch
                        {
                        }
                        finally
                        {
                            process.Dispose();
                        }
                    }
                }

                Thread.Sleep(400);
            }

            windowHandle = IntPtr.Zero;
            return false;
        }

        private void RestartSpecificOfficeApplication(string executableName, string friendlyName)
        {
            try
            {
                CloseRunningOfficeApplications();
                System.Threading.Thread.Sleep(180);
                string launchedTarget;
                if (!TryLaunchSpecificOfficeApplication(executableName, out launchedTarget))
                {
                    ShowErrorMessage("未能启动 " + friendlyName + "。请确认本机已安装对应 Office 组件。", "Office");
                }
            }
            catch (Exception ex)
            {
                ShowErrorMessage("重启 " + friendlyName + " 失败：" + ex.Message, "Office");
            }
        }

        private static OfficeCloseResult CloseRunningOfficeApplications()
        {
            var result = new OfficeCloseResult();
            var processNames = new[]
            {
                "WINWORD",
                "EXCEL",
                "POWERPNT",
                "OUTLOOK",
                "ONENOTE",
                "ONENOTEM",
                "VISIO",
                "MSACCESS",
                "MSPUB"
            };

            foreach (var processName in processNames)
            {
                Process[] processes;
                try
                {
                    processes = Process.GetProcessesByName(processName);
                }
                catch
                {
                    continue;
                }

                foreach (var process in processes)
                {
                    try
                    {
                        if (process.HasExited)
                        {
                            continue;
                        }

                        result.FoundCount++;
                        var closedGracefully = false;
                        try
                        {
                            if (process.MainWindowHandle != IntPtr.Zero)
                            {
                                closedGracefully = process.CloseMainWindow();
                            }
                        }
                        catch
                        {
                        }

                        if (closedGracefully)
                        {
                            try
                            {
                                if (process.WaitForExit(2500))
                                {
                                    result.GracefulCount++;
                                    continue;
                                }
                            }
                            catch
                            {
                            }
                        }

                        try
                        {
                            process.Kill();
                            if (process.WaitForExit(2500))
                            {
                                result.ForcedCount++;
                            }
                            else
                            {
                                result.FailedCount++;
                            }
                        }
                        catch
                        {
                            result.FailedCount++;
                        }
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }

            return result;
        }

        private static bool TryNormalizeOfficeProductKey(string rawValue, out string normalizedKey)
        {
            normalizedKey = string.Empty;
            var compact = Regex.Replace((rawValue ?? string.Empty).ToUpperInvariant(), @"[^A-Z0-9]", string.Empty);
            if (compact.Length != 25)
            {
                return false;
            }

            normalizedKey = string.Join("-", Enumerable.Range(0, 5).Select(index => compact.Substring(index * 5, 5)).ToArray());
            return true;
        }

        private static string NormalizeOfficeProductKey(string rawValue)
        {
            string normalizedKey;
            if (!TryNormalizeOfficeProductKey(rawValue, out normalizedKey))
            {
                throw new InvalidOperationException("读取到的 Office 密钥格式不正确，必须能还原成 25 位产品密钥。");
            }

            return normalizedKey;
        }

        private static string NormalizeLineEndings(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
        }

        private static string RemoveLooseOfficeKeyLines(string content)
        {
            var cleanedLines = new List<string>();
            foreach (var line in NormalizeLineEndings(content).Split('\n'))
            {
                string normalizedKey;
                if (TryNormalizeOfficeProductKey(line, out normalizedKey))
                {
                    continue;
                }

                cleanedLines.Add(line);
            }

            return string.Join("\n", cleanedLines.ToArray());
        }

        private static string TrimExcessBlankLines(string content)
        {
            var normalized = NormalizeLineEndings(content).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            normalized = Regex.Replace(normalized, @"\n{3,}", "\n\n");
            return normalized;
        }

        private static string BuildOfficeSection(string normalizedKey)
        {
            return OfficeKeySeparator + "\n\n" + OfficeKeyLabel + "\n\n" + normalizedKey + "\n\n" + OfficeKeySeparator;
        }

        private static string ReadFirstOfficeProductKey(string sourcePath, out List<string> sourceLines, out int keyLineIndex, out Encoding sourceEncoding)
        {
            if (!File.Exists(sourcePath))
            {
                throw new FileNotFoundException("未找到 Office 密钥源文件。", sourcePath);
            }

            sourceLines = ReadAllLinesWithOfficeEncoding(sourcePath, out sourceEncoding);
            keyLineIndex = -1;
            for (var i = 0; i < sourceLines.Count; i++)
            {
                var rawLine = sourceLines[i];
                string normalizedKey;
                if (TryNormalizeOfficeProductKey(rawLine, out normalizedKey))
                {
                    keyLineIndex = i;
                    return normalizedKey;
                }
            }

            throw new InvalidOperationException("密钥文件里没有找到有效的 25 位 Office 密钥，请补充。");
        }

        private static int CountAvailableOfficeProductKeys(string sourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            {
                return 0;
            }

            try
            {
                Encoding detectedEncoding;
                return ReadAllLinesWithOfficeEncoding(sourcePath, out detectedEncoding).Count(line =>
                {
                    string normalizedKey;
                    return TryNormalizeOfficeProductKey(line, out normalizedKey);
                });
            }
            catch
            {
                return 0;
            }
        }

        private static void RemoveOfficeProductKeyLine(string sourcePath, List<string> sourceLines, int keyLineIndex, Encoding sourceEncoding)
        {
            if (sourceLines == null || keyLineIndex < 0 || keyLineIndex >= sourceLines.Count)
            {
                return;
            }

            sourceLines.RemoveAt(keyLineIndex);
            WriteAllTextSafely(sourcePath, string.Join("\r\n", sourceLines.ToArray()), sourceEncoding ?? new UTF8Encoding(false));
        }

        private static void UpdateOfficeProductKeyFile(string targetPath, string key)
        {
            var normalizedKey = NormalizeOfficeProductKey(key);

            var directory = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            Encoding targetEncoding = null;
            var content = File.Exists(targetPath)
                ? ReadAllTextWithOfficeEncoding(targetPath, out targetEncoding)
                : string.Empty;
            var normalizedContent = NormalizeLineEndings(content);
            normalizedContent = OfficeProductKeyRegex.Replace(normalizedContent, string.Empty);
            normalizedContent = RemoveLooseOfficeKeyLines(normalizedContent);

            var titleMatch = OfficeKeyTitleRegex.Match(normalizedContent);
            if (titleMatch.Success)
            {
                var titleLineEnd = normalizedContent.IndexOf('\n', titleMatch.Index);
                if (titleLineEnd < 0)
                {
                    titleLineEnd = normalizedContent.Length;
                }

                var separatorMatch = OfficeKeySeparatorRegex.Match(normalizedContent, titleLineEnd);
                var beforeTitle = normalizedContent.Substring(0, titleLineEnd).TrimEnd();
                if (separatorMatch.Success)
                {
                    var afterSection = normalizedContent.Substring(separatorMatch.Index).TrimStart('\n');
                    normalizedContent = beforeTitle + "\n\n" + normalizedKey + "\n\n" + afterSection;
                }
                else
                {
                    normalizedContent = beforeTitle + "\n\n" + normalizedKey + "\n\n" + OfficeKeySeparator + "\n";
                }
            }
            else
            {
                var cleanBody = TrimExcessBlankLines(normalizedContent);
                normalizedContent = string.IsNullOrWhiteSpace(cleanBody)
                    ? BuildOfficeSection(normalizedKey)
                    : cleanBody + "\n\n" + BuildOfficeSection(normalizedKey);
            }

            WriteAllTextSafely(
                targetPath,
                TrimExcessBlankLines(normalizedContent).Replace("\n", "\r\n") + "\r\n",
                targetEncoding ?? new UTF8Encoding(false));
        }

        private static void WriteAllTextSafely(string path, string content, Encoding encoding)
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var tempFile = Path.Combine(
                string.IsNullOrWhiteSpace(directory) ? AppDomain.CurrentDomain.BaseDirectory : directory,
                Guid.NewGuid().ToString("N") + ".tmp");

            try
            {
                File.WriteAllText(tempFile, content ?? string.Empty, encoding ?? Encoding.UTF8);
                if (File.Exists(path))
                {
                    File.Replace(tempFile, path, null, true);
                }
                else
                {
                    File.Move(tempFile, path);
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }
                catch
                {
                }
            }
        }

        private void OpenOfficeActivation(Button sourceButton)
        {
            try
            {
                var sourcePath = BuildPortableCandidates(_settings.OfficeKeySourcePath, OfficeKeySourcePath).FirstOrDefault(File.Exists);
                var normalizedOperation = NormalizeOfficeKeyOperation(_settings.OfficeKeyOperation);
                var cutFirstLine = string.Equals(normalizedOperation, "cut", StringComparison.OrdinalIgnoreCase);
                var deleteFirstLine = string.Equals(normalizedOperation, "delete", StringComparison.OrdinalIgnoreCase);
                if (string.IsNullOrWhiteSpace(sourcePath))
                {
                    throw new FileNotFoundException("未找到 Office 密钥源文件。", string.IsNullOrWhiteSpace(_settings.OfficeKeySourcePath) ? OfficeKeySourcePath : _settings.OfficeKeySourcePath);
                }

                var availableKeys = CountAvailableOfficeProductKeys(sourcePath);
                if (availableKeys <= 0)
                {
                    UpdateOfficeActionButtonPresentation();
                    if (deleteFirstLine)
                    {
                        ShowInfo("密钥源文件当前没有可删除的有效密钥。", "Office");
                    }
                    else
                    {
                        ShowErrorMessage("密钥文件当前余 0 码，请先补充有效 Office 密钥。", "Office");
                    }
                    SyncOfficeActivationButtonStateAsync();
                    return;
                }

                List<string> sourceLines;
                int keyLineIndex;
                Encoding sourceEncoding;
                var officeKey = ReadFirstOfficeProductKey(sourcePath, out sourceLines, out keyLineIndex, out sourceEncoding);

                if (deleteFirstLine)
                {
                    RemoveOfficeProductKeyLine(sourcePath, sourceLines, keyLineIndex, sourceEncoding);
                    _officeDeleteCount++;
                    UpdateOfficeActionButtonPresentation();

                    var remainingDeleteKeys = CountAvailableOfficeProductKeys(sourcePath);
                    if (_settings.OfficeLowKeyWarningThreshold > 0 && remainingDeleteKeys < _settings.OfficeLowKeyWarningThreshold)
                    {
                        ShowInfo(
                            string.Format("提醒：密钥文件当前剩余 {0} 个有效密钥，已低于设置阈值 {1} 个，请及时补充。", remainingDeleteKeys, _settings.OfficeLowKeyWarningThreshold),
                            "Office 密钥提醒");
                    }

                    SyncOfficeActivationButtonStateAsync();
                    return;
                }

                var targetPath = SelectOfficeTargetPath();
                UpdateOfficeProductKeyFile(targetPath, officeKey);
                _officeDesktopWriteCount++;

                try
                {
                    CopyTextToClipboard(officeKey);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("密钥已写入桌面文件，但复制到剪贴板失败，所以后面无法直接粘贴。", ex);
                }

                if (cutFirstLine)
                {
                    RemoveOfficeProductKeyLine(sourcePath, sourceLines, keyLineIndex, sourceEncoding);
                    _officeCutCount++;
                }
                else
                {
                    _officeCopyCount++;
                }

                UpdateOfficeActionButtonPresentation();

                try
                {
                    CloseRunningOfficeApplications();
                    Thread.Sleep(180);
                    string launchedTarget;
                    if (!TryLaunchOfficeApplication(out launchedTarget))
                    {
                        ShowErrorMessage("密钥已处理，但未能自动打开 Office 软件，请确认本机已安装 Office。", "Office");
                    }
                }
                catch (Exception launchEx)
                {
                    ShowErrorMessage("密钥已处理，但打开 Office 软件失败：" + launchEx.Message, "Office");
                }

                var remainingKeys = CountAvailableOfficeProductKeys(sourcePath);
                if (_settings.OfficeLowKeyWarningThreshold > 0 && remainingKeys < _settings.OfficeLowKeyWarningThreshold)
                {
                    ShowInfo(
                        string.Format("提醒：密钥文件当前剩余 {0} 个有效密钥，已低于设置阈值 {1} 个，请及时补充。", remainingKeys, _settings.OfficeLowKeyWarningThreshold),
                        "Office 密钥提醒");
                }

                SyncOfficeActivationButtonStateAsync();
            }
            catch (Exception ex)
            {
                ShowErrorMessage(string.Format("Office 密钥处理失败：{0}", ex.Message), "Office");
                SyncOfficeActivationButtonStateAsync();
            }
        }

        private string FindSpeakerTestMedia()
        {
            return BuildPortableCandidates(
                _settings.SpeakerTestPath,
                @"资源文件\【俺妹】俺の妹がこんなに可愛いわけがないop「irony」.mp4").FirstOrDefault(File.Exists);
        }

        private void OpenSpeakerTest()
        {
            var speakerMedia = FindSpeakerTestMedia();
            if (speakerMedia == null)
            {
                ShowErrorMessage("未找到扬声器测试视频，请检查设置里的路径，或确认资源文件夹跟着软件一起放在 U 盘里。");
                return;
            }

            if (StartTarget(speakerMedia))
            {
                QueueRecentRecordCleanupForPath(speakerMedia);
            }
        }

        private void QueueRecentRecordCleanupForPath(string targetPath)
        {
            if (string.IsNullOrWhiteSpace(targetPath))
            {
                return;
            }

            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                var delays = new[] { 1200, 3200, 6200 };
                for (var i = 0; i < delays.Length; i++)
                {
                    System.Threading.Thread.Sleep(delays[i]);
                    try
                    {
                        DeleteRecentShortcutRecords(targetPath);
                        if (i == delays.Length - 1)
                        {
                            ClearExplorerHistory();
                        }
                    }
                    catch
                    {
                    }
                }
            });
        }

        private void DeleteRecentShortcutRecords(string targetPath)
        {
            try
            {
                var recentFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    @"Microsoft\Windows\Recent");
                if (!Directory.Exists(recentFolder))
                {
                    return;
                }

                var fileName = Path.GetFileName(targetPath) ?? string.Empty;
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(targetPath) ?? string.Empty;
                foreach (var shortcut in Directory.GetFiles(recentFolder, "*.lnk"))
                {
                    var shortcutName = Path.GetFileName(shortcut) ?? string.Empty;
                    if ((!string.IsNullOrWhiteSpace(fileName) && shortcutName.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0)
                        || (!string.IsNullOrWhiteSpace(fileNameWithoutExtension) && shortcutName.IndexOf(fileNameWithoutExtension, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        try
                        {
                            File.Delete(shortcut);
                        }
                        catch
                        {
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private int ClearExplorerHistory()
        {
            var failureCount = 0;
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var targets = new[]
            {
                Path.Combine(appData, @"Microsoft\Windows\Recent"),
                Path.Combine(appData, @"Microsoft\Windows\Recent\AutomaticDestinations"),
                Path.Combine(appData, @"Microsoft\Windows\Recent\CustomDestinations")
            };

            foreach (var folder in targets)
            {
                if (!Directory.Exists(folder))
                {
                    continue;
                }

                foreach (var file in Directory.GetFiles(folder))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                        failureCount++;
                    }
                }

                foreach (var directory in Directory.GetDirectories(folder))
                {
                    try
                    {
                        Directory.Delete(directory, true);
                    }
                    catch
                    {
                        failureCount++;
                    }
                }
            }

            var registryTargets = new[]
            {
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU",
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRU",
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU"
            };

            foreach (var target in registryTargets)
            {
                try
                {
                    Registry.CurrentUser.DeleteSubKeyTree(target, false);
                }
                catch
                {
                    failureCount++;
                }
            }

            return failureCount;
        }

        private int ClearConfiguredCleanupTargets()
        {
            var failureCount = 0;
            foreach (var rawPath in _settings.ExtraCleanupPaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var resolvedPath = ResolveConfiguredPath(rawPath);
                if (string.IsNullOrWhiteSpace(resolvedPath))
                {
                    continue;
                }

                try
                {
                    if (File.Exists(resolvedPath))
                    {
                        EnsurePathWritable(resolvedPath);
                        File.Delete(resolvedPath);
                        continue;
                    }

                    if (Directory.Exists(resolvedPath))
                    {
                        EnsureDirectoryTreeWritable(resolvedPath);
                        Directory.Delete(resolvedPath, true);
                    }
                }
                catch
                {
                    failureCount++;
                }
            }

            return failureCount;
        }

        private static void CloseRunningBrowserApplications()
        {
            var processNames = new[] { "msedge", "chrome", "brave", "firefox", "opera" };
            foreach (var processName in processNames)
            {
                Process[] processes;
                try
                {
                    processes = Process.GetProcessesByName(processName);
                }
                catch
                {
                    continue;
                }

                foreach (var process in processes)
                {
                    try
                    {
                        if (process.HasExited)
                        {
                            continue;
                        }

                        var closedGracefully = false;
                        try
                        {
                            if (process.MainWindowHandle != IntPtr.Zero)
                            {
                                closedGracefully = process.CloseMainWindow();
                            }
                        }
                        catch
                        {
                        }

                        if (closedGracefully)
                        {
                            try
                            {
                                if (process.WaitForExit(2200))
                                {
                                    continue;
                                }
                            }
                            catch
                            {
                            }
                        }

                        try
                        {
                            process.Kill();
                            process.WaitForExit(2200);
                        }
                        catch
                        {
                        }
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
        }

        private int ClearBrowserHistoryAndSessions()
        {
            var failureCount = 0;
            try
            {
                CloseRunningBrowserApplications();
                System.Threading.Thread.Sleep(500);
            }
            catch
            {
            }

            foreach (var profileRoot in GetChromiumBrowserProfileRoots())
            {
                failureCount += ClearChromiumBrowserData(profileRoot);
            }

            foreach (var profilePath in GetFirefoxProfilePaths())
            {
                failureCount += ClearFirefoxBrowserData(profilePath);
            }

            return failureCount;
        }

        private IEnumerable<string> GetChromiumBrowserProfileRoots()
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return new[]
            {
                Path.Combine(localAppData, @"Microsoft\Edge\User Data"),
                Path.Combine(localAppData, @"Google\Chrome\User Data"),
                Path.Combine(localAppData, @"BraveSoftware\Brave-Browser\User Data"),
                Path.Combine(localAppData, @"Opera Software\Opera Stable")
            }
            .Where(Directory.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        }

        private static IEnumerable<string> EnumerateChromiumProfileDirectories(string rootPath)
        {
            if (string.IsNullOrWhiteSpace(rootPath) || !Directory.Exists(rootPath))
            {
                return Enumerable.Empty<string>();
            }

            var roots = new List<string>();
            Action<string> addIfValid = candidate =>
            {
                if (string.IsNullOrWhiteSpace(candidate) || !Directory.Exists(candidate))
                {
                    return;
                }

                var name = Path.GetFileName(candidate) ?? string.Empty;
                if (string.Equals(name, "System Profile", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                if (!roots.Any(existing => string.Equals(existing, candidate, StringComparison.OrdinalIgnoreCase)))
                {
                    roots.Add(candidate);
                }
            };

            addIfValid(Path.Combine(rootPath, "Default"));
            foreach (var directory in Directory.GetDirectories(rootPath, "Profile *"))
            {
                addIfValid(directory);
            }

            var historyInRoot = Path.Combine(rootPath, "History");
            if (File.Exists(historyInRoot))
            {
                addIfValid(rootPath);
            }

            return roots;
        }

        private int ClearChromiumBrowserData(string userDataRoot)
        {
            var failureCount = 0;
            foreach (var profilePath in EnumerateChromiumProfileDirectories(userDataRoot))
            {
                var filesToDelete = new[]
                {
                    "History",
                    "History-journal",
                    "Visited Links",
                    "Top Sites",
                    "Top Sites-journal",
                    "Last Session",
                    "Last Tabs",
                    "Current Session",
                    "Current Tabs",
                    "Last Tabs",
                    "Last Session"
                };

                foreach (var fileName in filesToDelete)
                {
                    failureCount += DeleteFileIfExists(Path.Combine(profilePath, fileName));
                }

                var sessionsFolder = Path.Combine(profilePath, "Sessions");
                failureCount += DeleteDirectoryFiles(sessionsFolder);

                failureCount += NormalizeChromiumPreferences(Path.Combine(profilePath, "Preferences"));
            }

            failureCount += NormalizeChromiumPreferences(Path.Combine(userDataRoot, "Local State"));
            return failureCount;
        }

        private static int NormalizeChromiumPreferences(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return 0;
            }

            try
            {
                EnsurePathWritable(path);
                var content = File.ReadAllText(path, Encoding.UTF8);
                var updated = Regex.Replace(content, "\"exit_type\"\\s*:\\s*\"[^\"]*\"", "\"exit_type\":\"Normal\"", RegexOptions.IgnoreCase);
                updated = Regex.Replace(updated, "\"exited_cleanly\"\\s*:\\s*(false|0)", "\"exited_cleanly\":true", RegexOptions.IgnoreCase);
                if (!string.Equals(updated, content, StringComparison.Ordinal))
                {
                    WriteAllTextSafely(path, updated, Encoding.UTF8);
                }

                return 0;
            }
            catch
            {
                return 1;
            }
        }

        private IEnumerable<string> GetFirefoxProfilePaths()
        {
            var roamingProfiles = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Mozilla\Firefox\Profiles");
            if (!Directory.Exists(roamingProfiles))
            {
                return Enumerable.Empty<string>();
            }

            return Directory.GetDirectories(roamingProfiles)
                .Where(Directory.Exists)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private int ClearFirefoxBrowserData(string profilePath)
        {
            var failureCount = 0;
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "places.sqlite"));
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "places.sqlite-wal"));
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "places.sqlite-shm"));
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "sessionstore.jsonlz4"));
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "sessionCheckpoints.json"));
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "sessionstore-backups", "recovery.jsonlz4"));
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "sessionstore-backups", "recovery.baklz4"));
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "sessionstore-backups", "previous.jsonlz4"));
            failureCount += DeleteFileIfExists(Path.Combine(profilePath, "sessionstore-backups", "upgrade.jsonlz4"));
            failureCount += DeleteDirectoryFiles(Path.Combine(profilePath, "sessionstore-backups"));
            return failureCount;
        }

        private static int DeleteDirectoryFiles(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                return 0;
            }

            var failureCount = 0;
            try
            {
                foreach (var file in Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories))
                {
                    failureCount += DeleteFileIfExists(file);
                }
            }
            catch
            {
                failureCount++;
            }

            return failureCount;
        }

        private static int DeleteFileIfExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return 0;
            }

            try
            {
                EnsurePathWritable(path);
                File.Delete(path);
                return 0;
            }
            catch
            {
                return 1;
            }
        }

        private static void EnsurePathWritable(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return;
            }

            try
            {
                var attributes = File.GetAttributes(path);
                if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    File.SetAttributes(path, attributes & ~FileAttributes.ReadOnly);
                }
            }
            catch
            {
            }
        }

        private static void EnsureDirectoryTreeWritable(string directoryPath)
        {
            if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
            {
                return;
            }

            try
            {
                foreach (var file in Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories))
                {
                    EnsurePathWritable(file);
                }
            }
            catch
            {
            }
        }

        private void ClearHistoryAndRecycleBin(Button sourceButton)
        {
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                Exception cleanupError = null;
                var failureCount = 0;
                try
                {
                    failureCount = ClearExplorerHistory();
                    failureCount += ClearBrowserHistoryAndSessions();
                    failureCount += ClearConfiguredCleanupTargets();
                    var recycleResult = SHEmptyRecycleBin(IntPtr.Zero, null, ShrbNoconfirmation | ShrbNoprogressui | ShrbNosound);
                    if (!IsIgnorableRecycleBinResult(recycleResult))
                    {
                        throw new InvalidOperationException(string.Format("回收站清理失败，错误代码 {0}", recycleResult));
                    }

                    recycleResult = SHEmptyRecycleBin(IntPtr.Zero, null, ShrbNoconfirmation | ShrbNoprogressui | ShrbNosound);
                    if (!IsIgnorableRecycleBinResult(recycleResult))
                    {
                        throw new InvalidOperationException(string.Format("回收站清理失败，错误代码 {0}", recycleResult));
                    }
                }
                catch (Exception ex)
                {
                    cleanupError = ex;
                }

                BeginInvoke((MethodInvoker)delegate
                {
                    RefreshDashboard();

                    if (cleanupError == null && failureCount == 0)
                    {
                        if (sourceButton != null && !sourceButton.IsDisposed)
                        {
                            SetButtonState(sourceButton, "passed");
                        }
                    }
                    else
                    {
                        var message = cleanupError != null
                            ? string.Format("清理失败：{0}", cleanupError.Message)
                            : string.Format("清理未完全成功，仍有 {0} 项未能删除。", failureCount);
                        ShowErrorMessage(message);
                    }
                });
            });
        }

        private static bool IsIgnorableRecycleBinResult(uint recycleResult)
        {
            return recycleResult == 0 || recycleResult == 2147549183;
        }

        private void ShowPowerDialog()
        {
            using (var dialog = new Form())
            {
                dialog.Text = "重启或关机";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.ClientSize = new Size(330, 150);
                dialog.Font = Font;

                var label = new Label { Text = "请选择电源操作：", Location = new Point(24, 22), Size = new Size(180, 22) };
                var restartButton = new Button { Text = "立即重启", Location = new Point(24, 66), Size = new Size(82, 34) };
                var shutdownButton = new Button { Text = "立即关机", Location = new Point(124, 66), Size = new Size(82, 34) };
                var cancelButton = new Button { Text = "取消", Location = new Point(224, 66), Size = new Size(82, 34) };

                restartButton.Click += (_, __) =>
                {
                    if (MessageBox.Show(dialog, "确认现在重启吗？", "重启确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        StartTarget("shutdown.exe", "/r /t 0");
                        dialog.Close();
                    }
                };

                shutdownButton.Click += (_, __) =>
                {
                    if (MessageBox.Show(dialog, "确认现在关机吗？", "关机确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        StartTarget("shutdown.exe", "/s /t 0");
                        dialog.Close();
                    }
                };

                cancelButton.Click += (_, __) => dialog.Close();

                dialog.Controls.AddRange(new Control[]
                {
                    label, restartButton, shutdownButton, cancelButton
                });

                dialog.ShowDialog(this);
            }
        }
    }

    internal sealed class WifiProfile
    {
        public string Ssid = string.Empty;
        public string Password = string.Empty;
        public string Authentication = "WPA2PSK";
        public string Encryption = "AES";
    }

    internal sealed class ButtonStateInfo
    {
        public string BaseTitle = string.Empty;
        public string Subtitle = string.Empty;
        public string State = "default";
        public bool TrackCompletion;
        public Button LinkedButton;
        public bool CompactModeStyle;
    }

    internal enum StartupMode
    {
        None,
        Zen,
        Speedrun
    }

    internal enum DisplayMode
    {
        Normal,
        Zen,
        Speedrun
    }

    internal enum WlanIntfOpcode
    {
        AutoconfEnabled = 1,
        BackgroundScanEnabled = 2,
        MediaStreamingMode = 3,
        RadioState = 4
    }

    internal enum Dot11RadioState : uint
    {
        Unknown = 0,
        On = 1,
        Off = 2
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct WlanPhyRadioState
    {
        public uint PhyIndex;
        public Dot11RadioState SoftwareRadioState;
        public Dot11RadioState HardwareRadioState;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct WlanRadioState
    {
        public uint NumberOfPhys;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)]
        public WlanPhyRadioState[] PhyRadioState;
    }

    internal sealed class WirelessEnvironmentStatus
    {
        public bool HasWirelessAdapter;
        public bool HasEnabledAdapter;
        public bool HasInterface;
        public bool SoftwareRadioOff;
        public bool HardwareRadioOff;
        public bool CanAttemptConnect;
        public string CommandError = string.Empty;
        public string Message = string.Empty;
        public List<string> InterfaceNames = new List<string>();
    }

    internal sealed class OfficeCloseResult
    {
        public int FoundCount;
        public int GracefulCount;
        public int ForcedCount;
        public int FailedCount;
    }

    internal sealed class OfficePostActionStep
    {
        public int DelayMs;
        public string Keys = string.Empty;
        public byte VirtualKeyCode;
    }

    internal sealed class OfficePostActionRule
    {
        public string Condition = string.Empty;
        public string Sequence = string.Empty;
    }

    internal enum OfficePostTriggerKind
    {
        Immediate,
        Delay,
        Key,
        PopupClosed
    }

    internal sealed class OfficePostTriggerState
    {
        public OfficePostActionRule Rule = new OfficePostActionRule();
        public OfficePostTriggerKind Kind;
        public Keys TriggerKey = Keys.None;
        public int DelayMs;
        public bool PreviousDown;
        public Stopwatch StartedAt;
    }

    internal sealed class AppSettings
    {
        public bool TopMost;
        public readonly List<WifiProfile> WifiProfiles = new List<WifiProfile>();
        public int BatteryGreenHealthPercent = 80;
        public int BatteryRedHealthPercent = 60;
        public int BatteryGreenCapacityMWh;
        public int BatteryRedCapacityMWh;
        public string SoftwareDisplayName = string.Empty;
        public string OfficeKeySourcePath = @"3.txt";
        public readonly List<string> OfficeKeyTargetPaths = new List<string>();
        public string OfficeKeyTargetPath = string.Empty;
        public string OfficeKeyOperation = "copy";
        public int OfficeLowKeyWarningThreshold;
        public readonly List<string> OfficeAppPaths = new List<string>();
        public string OfficeAppPath = string.Empty;
        public string OfficeCommunicationTestUrl = "http://10.229.159.132:8000";
        public string OfficeCommunicationTestPayload = "Office 通信测试\r\n只发送模拟文本，不包含激活数据。";
        public int OfficeActivationCount;
        public bool OfficePostActionEnabled = true;
        public string OfficePostActionCondition = "关闭激活弹窗后";
        public string OfficePostActionSequence =
            "按键:{Alt}\r\n" +
            "停顿:250\r\n" +
            "按键:D\r\n" +
            "停顿:400\r\n" +
            "按键:Y\r\n" +
            "停顿:200\r\n" +
            "按键:4\r\n" +
            "停顿:500\r\n" +
            "按键:{Tab}\r\n" +
            "停顿:250\r\n" +
            "按键:{Tab}\r\n" +
            "停顿:250\r\n" +
            "按键:{Tab}\r\n" +
            "停顿:500\r\n" +
            "按键:^v\r\n" +
            "停顿:500\r\n" +
            "按键:{Tab}\r\n" +
            "停顿:500\r\n" +
            "按键:{Enter}";
        public bool OfficePostAhkEnabled;
        public string OfficePostAhkPath = "AutoHotkey.exe";
        public string OfficePostAhkScript = string.Empty;
        public string CameraCheckPath = string.Empty;
        public string KeyboardCheckPath = @"Keyboard Test Utility.exe";
        public string SpeakerTestPath = @"资源文件\【俺妹】俺の妹がこんなに可愛いわけがないop「irony」.mp4";
        public string IconFilePath = string.Empty;
        public string WindowsActivationCopyTemplate = "slmgr /ipk {key}";
        public readonly List<string> ExtraCleanupPaths = new List<string>();
        public string GithubUrl = string.Empty;
        public string DonateUrl = string.Empty;
        public string SoftwareUpdatedAt = "2026-03-21";
        public string ThemeMode = "light";
        public int WindowWidth;
        public int WindowHeight;
        public int WindowLeft = int.MinValue;
        public int WindowTop = int.MinValue;
        public bool WindowMaximized;
        public int ZenWindowLeft = int.MinValue;
        public int ZenWindowTop = int.MinValue;
        public int SpeedrunWindowLeft = int.MinValue;
        public int SpeedrunWindowTop = int.MinValue;

        private static string SettingsPath
        {
            get
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                return Path.Combine(baseDir, "settings_v13.ini");
            }
        }

        private static string LegacySettingsPath
        {
            get
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                return Path.Combine(baseDir, "settings.ini");
            }
        }

        public static AppSettings Load()
        {
            var settings = new AppSettings();
            var sourceSettingsPath = File.Exists(SettingsPath) ? SettingsPath : LegacySettingsPath;
            if (!File.Exists(sourceSettingsPath))
            {
                return settings;
            }

            foreach (var rawLine in File.ReadAllLines(sourceSettingsPath, Encoding.UTF8))
            {
                var line = rawLine.Trim();
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                {
                    continue;
                }

                if (line.StartsWith("topmost=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.TopMost = string.Equals(line.Substring(8), "true", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (line.StartsWith("wifi=", StringComparison.OrdinalIgnoreCase))
                {
                    var parts = line.Substring(5).Split(new[] { '|' }, 4);
                    if (parts.Length >= 4)
                    {
                        settings.WifiProfiles.Add(new WifiProfile
                        {
                            Ssid = Unescape(parts[0]),
                            Password = Unescape(parts[1]),
                            Authentication = Unescape(parts[2]),
                            Encryption = Unescape(parts[3])
                        });
                    }

                    continue;
                }

                if (line.StartsWith("battery_green_health=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("battery_green_health=".Length), out value))
                    {
                        settings.BatteryGreenHealthPercent = value;
                    }
                    continue;
                }

                if (line.StartsWith("battery_red_health=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("battery_red_health=".Length), out value))
                    {
                        settings.BatteryRedHealthPercent = value;
                    }
                    continue;
                }

                if (line.StartsWith("battery_green_capacity=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("battery_green_capacity=".Length), out value))
                    {
                        settings.BatteryGreenCapacityMWh = value;
                    }
                    continue;
                }

                if (line.StartsWith("battery_red_capacity=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("battery_red_capacity=".Length), out value))
                    {
                        settings.BatteryRedCapacityMWh = value;
                    }

                    continue;
                }

                if (line.StartsWith("software_display_name=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.SoftwareDisplayName = Unescape(line.Substring("software_display_name=".Length));
                    continue;
                }

                if (line.StartsWith("office_key_source=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficeKeySourcePath = Unescape(line.Substring("office_key_source=".Length));
                    continue;
                }

                if (line.StartsWith("office_key_target=", StringComparison.OrdinalIgnoreCase))
                {
                    var officeTargetPath = Unescape(line.Substring("office_key_target=".Length));
                    settings.OfficeKeyTargetPath = officeTargetPath;
                    if (!string.IsNullOrWhiteSpace(officeTargetPath)
                        && !settings.OfficeKeyTargetPaths.Any(existing => string.Equals(existing, officeTargetPath, StringComparison.OrdinalIgnoreCase)))
                    {
                        settings.OfficeKeyTargetPaths.Add(officeTargetPath);
                    }
                    continue;
                }

                if (line.StartsWith("office_key_operation=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficeKeyOperation = Unescape(line.Substring("office_key_operation=".Length));
                    continue;
                }

                if (line.StartsWith("office_low_key_warning=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("office_low_key_warning=".Length), out value))
                    {
                        settings.OfficeLowKeyWarningThreshold = Math.Max(0, value);
                    }
                    continue;
                }

                if (line.StartsWith("office_app_path=", StringComparison.OrdinalIgnoreCase))
                {
                    var officeAppPath = Unescape(line.Substring("office_app_path=".Length));
                    settings.OfficeAppPath = officeAppPath;
                    if (!string.IsNullOrWhiteSpace(officeAppPath)
                        && !settings.OfficeAppPaths.Any(existing => string.Equals(existing, officeAppPath, StringComparison.OrdinalIgnoreCase)))
                    {
                        settings.OfficeAppPaths.Add(officeAppPath);
                    }
                    continue;
                }

                if (line.StartsWith("office_activation_count=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("office_activation_count=".Length), out value))
                    {
                        settings.OfficeActivationCount = Math.Max(0, value);
                    }
                    continue;
                }

                if (line.StartsWith("office_post_action_enabled=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficePostActionEnabled = string.Equals(line.Substring("office_post_action_enabled=".Length), "true", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (line.StartsWith("office_post_action_condition=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficePostActionCondition = Unescape(line.Substring("office_post_action_condition=".Length));
                    continue;
                }

                if (line.StartsWith("office_post_action_sequence=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficePostActionSequence = Unescape(line.Substring("office_post_action_sequence=".Length));
                    continue;
                }

                if (line.StartsWith("office_post_ahk_enabled=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficePostAhkEnabled = string.Equals(line.Substring("office_post_ahk_enabled=".Length), "true", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (line.StartsWith("office_post_ahk_path=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficePostAhkPath = Unescape(line.Substring("office_post_ahk_path=".Length));
                    continue;
                }

                if (line.StartsWith("office_post_ahk_script=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficePostAhkScript = Unescape(line.Substring("office_post_ahk_script=".Length));
                    continue;
                }

                if (line.StartsWith("office_comm_test_url=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficeCommunicationTestUrl = Unescape(line.Substring("office_comm_test_url=".Length));
                    continue;
                }

                if (line.StartsWith("office_comm_test_payload=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.OfficeCommunicationTestPayload = Unescape(line.Substring("office_comm_test_payload=".Length));
                    continue;
                }

                if (line.StartsWith("camera_check_path=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.CameraCheckPath = Unescape(line.Substring("camera_check_path=".Length));
                    continue;
                }

                if (line.StartsWith("keyboard_check_path=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.KeyboardCheckPath = Unescape(line.Substring("keyboard_check_path=".Length));
                    continue;
                }

                if (line.StartsWith("speaker_test_path=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.SpeakerTestPath = Unescape(line.Substring("speaker_test_path=".Length));
                    continue;
                }

                if (line.StartsWith("icon_file_path=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.IconFilePath = Unescape(line.Substring("icon_file_path=".Length));
                    continue;
                }

                if (line.StartsWith("windows_activation_copy_template=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.WindowsActivationCopyTemplate = Unescape(line.Substring("windows_activation_copy_template=".Length));
                    continue;
                }

                if (line.StartsWith("extra_cleanup_path=", StringComparison.OrdinalIgnoreCase))
                {
                    var cleanupPath = Unescape(line.Substring("extra_cleanup_path=".Length));
                    if (!string.IsNullOrWhiteSpace(cleanupPath)
                        && !settings.ExtraCleanupPaths.Any(existing => string.Equals(existing, cleanupPath, StringComparison.OrdinalIgnoreCase)))
                    {
                        settings.ExtraCleanupPaths.Add(cleanupPath);
                    }
                    continue;
                }

                if (line.StartsWith("github_url=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.GithubUrl = Unescape(line.Substring("github_url=".Length));
                    continue;
                }

                if (line.StartsWith("donate_url=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.DonateUrl = Unescape(line.Substring("donate_url=".Length));
                    continue;
                }

                if (line.StartsWith("software_updated_at=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.SoftwareUpdatedAt = Unescape(line.Substring("software_updated_at=".Length));
                    continue;
                }

                if (line.StartsWith("theme_mode=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.ThemeMode = Unescape(line.Substring("theme_mode=".Length));
                    continue;
                }

                if (line.StartsWith("window_width=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("window_width=".Length), out value))
                    {
                        settings.WindowWidth = value;
                    }
                    continue;
                }

                if (line.StartsWith("window_height=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("window_height=".Length), out value))
                    {
                        settings.WindowHeight = value;
                    }
                    continue;
                }

                if (line.StartsWith("window_left=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("window_left=".Length), out value))
                    {
                        settings.WindowLeft = value;
                    }
                    continue;
                }

                if (line.StartsWith("window_top=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("window_top=".Length), out value))
                    {
                        settings.WindowTop = value;
                    }
                    continue;
                }

                if (line.StartsWith("window_maximized=", StringComparison.OrdinalIgnoreCase))
                {
                    settings.WindowMaximized = string.Equals(line.Substring("window_maximized=".Length), "true", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (line.StartsWith("zen_window_left=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("zen_window_left=".Length), out value))
                    {
                        settings.ZenWindowLeft = value;
                    }
                    continue;
                }

                if (line.StartsWith("zen_window_top=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("zen_window_top=".Length), out value))
                    {
                        settings.ZenWindowTop = value;
                    }
                    continue;
                }

                if (line.StartsWith("speedrun_window_left=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("speedrun_window_left=".Length), out value))
                    {
                        settings.SpeedrunWindowLeft = value;
                    }
                    continue;
                }

                if (line.StartsWith("speedrun_window_top=", StringComparison.OrdinalIgnoreCase))
                {
                    int value;
                    if (int.TryParse(line.Substring("speedrun_window_top=".Length), out value))
                    {
                        settings.SpeedrunWindowTop = value;
                    }
                }
            }

            if (ShouldUpgradeOfficePostActionSequence(settings.OfficePostActionSequence))
            {
                settings.OfficePostActionSequence = new AppSettings().OfficePostActionSequence;
            }

            return settings;
        }

        private static bool ShouldUpgradeOfficePostActionSequence(string sequence)
        {
            var normalized = (sequence ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return true;
            }

            return string.Equals(normalized, "Alt--D--Y4--tab*3--ctrl+V--tab--回车", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "{Alt}--D--Sleep 500--Y4--Sleep 500--{Tab 3}--^v--{Tab}--{Enter}", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "{Alt}--Sleep 250--D--Sleep 400--Y4--Sleep 500--{Tab}--Sleep 250--{Tab}--Sleep 250--{Tab}--Sleep 500--^v--Sleep 500--{Tab}--Sleep 500--{Enter}", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "{Alt}--Sleep 250--D--Sleep 400--Y--Sleep 200--4--Sleep 500--{Tab}--Sleep 250--{Tab}--Sleep 250--{Tab}--Sleep 500--^v--Sleep 500--{Tab}--Sleep 500--{Enter}", StringComparison.OrdinalIgnoreCase);
        }

        public void Save()
        {
            var lines = new List<string>();
            lines.Add("# PcCheckTool settings");
            lines.Add("topmost=" + (TopMost ? "true" : "false"));
            lines.Add("battery_green_health=" + BatteryGreenHealthPercent);
            lines.Add("battery_red_health=" + BatteryRedHealthPercent);
            lines.Add("battery_green_capacity=" + BatteryGreenCapacityMWh);
            lines.Add("battery_red_capacity=" + BatteryRedCapacityMWh);
            lines.Add("software_display_name=" + Escape(SoftwareDisplayName));
            lines.Add("office_key_source=" + Escape(OfficeKeySourcePath));
            var officeTargetPaths = OfficeKeyTargetPaths
                .Select(path => (path ?? string.Empty).Trim())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (officeTargetPaths.Count == 0 && !string.IsNullOrWhiteSpace(OfficeKeyTargetPath))
            {
                officeTargetPaths.Add(OfficeKeyTargetPath.Trim());
            }
            if (officeTargetPaths.Count == 0)
            {
                officeTargetPaths.Add(string.Empty);
            }
            OfficeKeyTargetPath = officeTargetPaths[0];
            foreach (var officeTargetPath in officeTargetPaths)
            {
                lines.Add("office_key_target=" + Escape(officeTargetPath));
            }
            lines.Add("office_key_operation=" + Escape(OfficeKeyOperation));
            lines.Add("office_low_key_warning=" + OfficeLowKeyWarningThreshold);
            var officeAppPaths = OfficeAppPaths
                .Select(path => (path ?? string.Empty).Trim())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (officeAppPaths.Count == 0 && !string.IsNullOrWhiteSpace(OfficeAppPath))
            {
                officeAppPaths.Add(OfficeAppPath.Trim());
            }
            if (officeAppPaths.Count > 0)
            {
                OfficeAppPath = officeAppPaths[0];
            }
            foreach (var officeAppPath in officeAppPaths)
            {
                lines.Add("office_app_path=" + Escape(officeAppPath));
            }
            lines.Add("office_activation_count=" + OfficeActivationCount);
            lines.Add("office_post_action_enabled=" + (OfficePostActionEnabled ? "true" : "false"));
            lines.Add("office_post_action_condition=" + Escape(OfficePostActionCondition));
            lines.Add("office_post_action_sequence=" + Escape(OfficePostActionSequence));
            lines.Add("office_post_ahk_enabled=" + (OfficePostAhkEnabled ? "true" : "false"));
            lines.Add("office_post_ahk_path=" + Escape(OfficePostAhkPath));
            lines.Add("office_post_ahk_script=" + Escape(OfficePostAhkScript));
            lines.Add("office_comm_test_url=" + Escape(OfficeCommunicationTestUrl));
            lines.Add("office_comm_test_payload=" + Escape(OfficeCommunicationTestPayload));
            lines.Add("camera_check_path=" + Escape(CameraCheckPath));
            lines.Add("keyboard_check_path=" + Escape(KeyboardCheckPath));
            lines.Add("speaker_test_path=" + Escape(SpeakerTestPath));
            lines.Add("icon_file_path=" + Escape(IconFilePath));
            lines.Add("windows_activation_copy_template=" + Escape(WindowsActivationCopyTemplate));
            foreach (var cleanupPath in ExtraCleanupPaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                lines.Add("extra_cleanup_path=" + Escape(cleanupPath));
            }
            lines.Add("github_url=" + Escape(GithubUrl));
            lines.Add("donate_url=" + Escape(DonateUrl));
            lines.Add("software_updated_at=" + Escape(SoftwareUpdatedAt));
            lines.Add("theme_mode=" + Escape(ThemeMode));
            lines.Add("window_width=" + WindowWidth);
            lines.Add("window_height=" + WindowHeight);
            lines.Add("window_left=" + WindowLeft);
            lines.Add("window_top=" + WindowTop);
            lines.Add("window_maximized=" + (WindowMaximized ? "true" : "false"));
            lines.Add("zen_window_left=" + ZenWindowLeft);
            lines.Add("zen_window_top=" + ZenWindowTop);
            lines.Add("speedrun_window_left=" + SpeedrunWindowLeft);
            lines.Add("speedrun_window_top=" + SpeedrunWindowTop);

            foreach (var profile in WifiProfiles)
            {
                lines.Add(string.Format(
                    "wifi={0}|{1}|{2}|{3}",
                    Escape(profile.Ssid),
                    Escape(profile.Password),
                    Escape(profile.Authentication),
                    Escape(profile.Encryption)));
            }

            var tempPath = SettingsPath + ".tmp";
            File.WriteAllLines(tempPath, lines.ToArray(), Encoding.UTF8);
            if (File.Exists(SettingsPath))
            {
                File.Replace(tempPath, SettingsPath, null, true);
            }
            else
            {
                File.Move(tempPath, SettingsPath);
            }
        }

        private static string Escape(string value)
        {
            return (value ?? string.Empty)
                .Replace("\r\n", "<<NL>>")
                .Replace("\n", "<<NL>>")
                .Replace("\\", "\\\\")
                .Replace("|", "\\p");
        }

        private static string Unescape(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\p", "|")
                .Replace("\\\\", "\\")
                .Replace("<<NL>>", Environment.NewLine);
        }
    }
}
