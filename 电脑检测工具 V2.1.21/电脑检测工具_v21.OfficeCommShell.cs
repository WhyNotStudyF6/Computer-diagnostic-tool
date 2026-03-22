using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Management;
using System.Threading;
using System.Windows.Forms;

namespace PcCheckTool
{
    internal sealed partial class MainForm
    {
        private string ResolveOfficeCommunicationExecutionDirectory(string configuredPath)
        {
            var candidates = BuildPortableCandidates(configuredPath);
            foreach (var candidate in candidates)
            {
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    continue;
                }

                if (Directory.Exists(candidate))
                {
                    return candidate;
                }

                if (File.Exists(candidate))
                {
                    return Path.GetDirectoryName(candidate) ?? string.Empty;
                }
            }

            var resolved = ResolveConfiguredPath(configuredPath);
            if (Directory.Exists(resolved))
            {
                return resolved;
            }

            if (File.Exists(resolved))
            {
                return Path.GetDirectoryName(resolved) ?? string.Empty;
            }

            return string.Empty;
        }

        private List<string> GetOfficeCommunicationExecutablePaths(string configuredPath)
        {
            var executionDirectory = ResolveOfficeCommunicationExecutionDirectory(configuredPath);
            if (string.IsNullOrWhiteSpace(executionDirectory) || !Directory.Exists(executionDirectory))
            {
                return new List<string>();
            }

            var allowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".exe",
                ".bat",
                ".cmd",
                ".ps1"
            };

            return Directory.GetFiles(executionDirectory)
                .Where(path => allowedExtensions.Contains(Path.GetExtension(path) ?? string.Empty))
                .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private ProcessStartInfo BuildOfficeCommunicationProcessStartInfo(string executablePath)
        {
            var extension = (Path.GetExtension(executablePath) ?? string.Empty).ToLowerInvariant();
            var workingDirectory = Path.GetDirectoryName(executablePath) ?? GetApplicationDirectory();
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = workingDirectory
            };

            switch (extension)
            {
                case ".bat":
                case ".cmd":
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = "/c \"" + executablePath + "\"";
                    break;
                case ".ps1":
                    startInfo.FileName = "powershell.exe";
                    startInfo.Arguments = "-ExecutionPolicy Bypass -File \"" + executablePath + "\"";
                    break;
                default:
                    startInfo.FileName = executablePath;
                    break;
            }

            return startInfo;
        }

        private static IEnumerable<int> GetChildProcessIds(int parentProcessId)
        {
            var result = new List<int>();
            if (parentProcessId <= 0)
            {
                return result;
            }

            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    string.Format("SELECT ProcessId FROM Win32_Process WHERE ParentProcessId = {0}", parentProcessId)))
                {
                    foreach (var item in searcher.Get().Cast<ManagementObject>())
                    {
                        try
                        {
                            result.Add(Convert.ToInt32(item["ProcessId"]));
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

            return result;
        }

        private static HashSet<int> GetProcessTreeIds(int rootProcessId)
        {
            var result = new HashSet<int>();
            if (rootProcessId <= 0)
            {
                return result;
            }

            var queue = new Queue<int>();
            queue.Enqueue(rootProcessId);
            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                foreach (var childProcessId in GetChildProcessIds(current))
                {
                    if (result.Add(childProcessId))
                    {
                        queue.Enqueue(childProcessId);
                    }
                }
            }

            return result;
        }

        private static bool IsProcessRunning(int processId)
        {
            if (processId <= 0)
            {
                return false;
            }

            try
            {
                using (var process = Process.GetProcessById(processId))
                {
                    return !process.HasExited;
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool IsProcessTreeRunning(int rootProcessId)
        {
            if (rootProcessId <= 0)
            {
                return false;
            }

            if (IsProcessRunning(rootProcessId))
            {
                return true;
            }

            foreach (var childProcessId in GetProcessTreeIds(rootProcessId))
            {
                if (IsProcessRunning(childProcessId))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool WaitForProcessTreeExit(int rootProcessId, Func<bool> cancellationCheck, Action onTick)
        {
            while (IsProcessTreeRunning(rootProcessId))
            {
                if (cancellationCheck != null && cancellationCheck())
                {
                    TerminateProcessTree(rootProcessId);
                    return false;
                }

                Thread.Sleep(120);
                if (onTick != null)
                {
                    onTick();
                }
            }

            return true;
        }

        private static void TerminateProcessTree(int rootProcessId)
        {
            if (rootProcessId <= 0)
            {
                return;
            }

            try
            {
                using (var killer = new Process())
                {
                    killer.StartInfo = new ProcessStartInfo
                    {
                        FileName = "taskkill.exe",
                        Arguments = string.Format("/PID {0} /T /F", rootProcessId),
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    killer.Start();
                    killer.WaitForExit(5000);
                }
            }
            catch
            {
                try
                {
                    using (var process = Process.GetProcessById(rootProcessId))
                    {
                        if (!process.HasExited)
                        {
                            process.Kill();
                        }
                    }
                }
                catch
                {
                }
            }
        }

        private static string FormatCommunicationElapsed(TimeSpan elapsed)
        {
            if (elapsed.TotalMinutes >= 1)
            {
                return string.Format("{0:mm\\:ss\\.ff}", elapsed);
            }

            return string.Format("{0:0.00}s", elapsed.TotalSeconds);
        }

        private static void StyleTaskFlowSurface(Panel panel)
        {
            panel.BackColor = Color.White;
            panel.Paint += (_, e) =>
            {
                if (e == null)
                {
                    return;
                }

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var borderPen = new Pen(Color.FromArgb(226, 232, 240), 1f))
                {
                    var rect = new Rectangle(0, 0, panel.Width - 1, panel.Height - 1);
                    e.Graphics.DrawRectangle(borderPen, rect);
                }
            };
        }

        private static Panel CreateTaskFlowConnector()
        {
            var connector = new Panel
            {
                Width = 34,
                Height = 58,
                Margin = new Padding(2, 10, 2, 10),
                BackColor = Color.White
            };
            connector.Paint += (_, e) =>
            {
                if (e == null)
                {
                    return;
                }

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var linePen = new Pen(Color.FromArgb(191, 219, 254), 2f))
                {
                    var midY = connector.Height / 2;
                    e.Graphics.DrawLine(linePen, 4, midY, connector.Width - 10, midY);
                    e.Graphics.DrawLine(linePen, connector.Width - 14, midY - 4, connector.Width - 8, midY);
                    e.Graphics.DrawLine(linePen, connector.Width - 14, midY + 4, connector.Width - 8, midY);
                }
            };
            return connector;
        }

        private static void StyleTaskFlowButton(Button button, bool primary)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = primary ? 0 : 1;
            button.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            button.BackColor = primary ? Color.FromArgb(37, 99, 235) : Color.White;
            button.ForeColor = primary ? Color.White : Color.FromArgb(30, 41, 59);
            button.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            button.Cursor = Cursors.Hand;
        }

        private static void StyleTaskFlowInput(TextBoxBase input, bool strong = false)
        {
            input.BackColor = Color.White;
            input.ForeColor = Color.FromArgb(15, 23, 42);
            input.BorderStyle = BorderStyle.FixedSingle;
            if (strong)
            {
                input.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            }
        }

        private void ShowOfficeCommunicationShellDialog()
        {
            using (var dialog = new Form())
            {
                dialog.Text = string.IsNullOrWhiteSpace(_settings.OfficeCommunicationWindowTitle)
                    ? DefaultOfficeCommunicationWindowTitle
                    : _settings.OfficeCommunicationWindowTitle;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.AutoScaleMode = AutoScaleMode.Dpi;
                dialog.KeyPreview = true;
                dialog.ClientSize = new Size(860, 674);
                dialog.Font = Font;
                dialog.BackColor = Color.FromArgb(244, 247, 251);

                var topCard = new Panel { Location = new Point(18, 14), Size = new Size(822, 302) };
                var stepCard = new Panel { Location = new Point(18, 328), Size = new Size(822, 148) };
                var outputCard = new Panel { Location = new Point(18, 488), Size = new Size(822, 132) };
                StyleTaskFlowSurface(topCard);
                StyleTaskFlowSurface(stepCard);
                StyleTaskFlowSurface(outputCard);

                var titleBox = new TextBox
                {
                    Location = new Point(18, 16),
                    Size = new Size(252, 32),
                    Text = dialog.Text,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point)
                };
                StyleTaskFlowInput(titleBox, true);
                titleBox.TextChanged += (_, __) =>
                {
                    dialog.Text = string.IsNullOrWhiteSpace(titleBox.Text)
                        ? DefaultOfficeCommunicationWindowTitle
                        : titleBox.Text.Trim();
                };

                var escTipLink = new LinkLabel
                {
                    Text = "按 Esc 退出窗口",
                    Location = new Point(286, 20),
                    Size = new Size(130, 24),
                    LinkColor = Color.FromArgb(37, 99, 235),
                    ActiveLinkColor = Color.FromArgb(29, 78, 216),
                    VisitedLinkColor = Color.FromArgb(37, 99, 235)
                };
                var totalElapsedLabel = new Label
                {
                    Text = "总耗时：0.00s",
                    Location = new Point(620, 20),
                    Size = new Size(184, 22),
                    TextAlign = ContentAlignment.MiddleRight,
                    ForeColor = Color.DimGray
                };
                totalElapsedLabel.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);

                var descriptionLabel = new Label { Text = "窗口说明", Location = new Point(18, 62), Size = new Size(120, 22), BackColor = Color.White };
                var descriptionBox = new TextBox
                {
                    Location = new Point(18, 88),
                    Size = new Size(786, 52),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true,
                    Text = string.IsNullOrWhiteSpace(_settings.OfficeCommunicationWindowDescription)
                        ? DefaultOfficeCommunicationWindowDescription
                        : _settings.OfficeCommunicationWindowDescription
                };
                StyleTaskFlowInput(descriptionBox);
                var directoryLabel = new Label { Text = "执行目录", Location = new Point(18, 150), Size = new Size(120, 22), BackColor = Color.White };
                var directoryBox = new TextBox
                {
                    Location = new Point(18, 176),
                    Size = new Size(786, 30),
                    Text = _settings.OfficeCommunicationTestUrl ?? string.Empty
                };
                StyleTaskFlowInput(directoryBox);
                var noteLabel = new Label { Text = "附加文字/备注", Location = new Point(18, 216), Size = new Size(160, 22), BackColor = Color.White };
                var noteBox = new TextBox
                {
                    Location = new Point(18, 242),
                    Size = new Size(786, 42),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true,
                    Text = _settings.OfficeCommunicationTestPayload ?? string.Empty
                };
                StyleTaskFlowInput(noteBox);
                var summaryLabel = new Label
                {
                    Text = "等待执行",
                    Location = new Point(16, 14),
                    Size = new Size(790, 22),
                    ForeColor = Color.DimGray,
                    BackColor = Color.White
                };
                summaryLabel.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point);
                var stepFlow = new FlowLayoutPanel
                {
                    Location = new Point(16, 42),
                    Size = new Size(790, 90),
                    WrapContents = false,
                    AutoScroll = true,
                    FlowDirection = FlowDirection.LeftToRight,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.White,
                    Padding = new Padding(2)
                };
                var outputLabel = new Label
                {
                    Text = "执行输出/结果(可复制)",
                    Location = new Point(16, 14),
                    Size = new Size(200, 22),
                    BackColor = Color.White
                };
                var outputBox = new RichTextBox
                {
                    Location = new Point(16, 40),
                    Size = new Size(790, 76),
                    ReadOnly = true,
                    ShortcutsEnabled = true,
                    DetectUrls = false,
                    BorderStyle = BorderStyle.FixedSingle,
                    Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
                };
                StyleTaskFlowInput(outputBox);
                var rerunButton = new Button { Text = "重新执行", Location = new Point(18, 630), Size = new Size(96, 32) };
                var saveDefaultsButton = new Button { Text = "保存为默认", Location = new Point(120, 630), Size = new Size(104, 32) };
                var copyAllButton = new Button { Text = "复制全部", Location = new Point(230, 630), Size = new Size(96, 32) };
                var closeButton = new Button { Text = "关闭", Location = new Point(744, 630), Size = new Size(96, 32) };
                StyleTaskFlowButton(rerunButton, true);
                StyleTaskFlowButton(saveDefaultsButton, false);
                StyleTaskFlowButton(copyAllButton, false);
                StyleTaskFlowButton(closeButton, false);

                var stepIcons = new List<Label>();
                var stepTimes = new List<Label>();
                var cancellationRequested = false;
                var isRunning = false;
                Process currentProcess = null;
                var currentProcessId = 0;
                var totalStopwatch = new Stopwatch();

                Action<Action> ui = action =>
                {
                    try
                    {
                        if (!dialog.IsDisposed && dialog.IsHandleCreated)
                        {
                            dialog.BeginInvoke((MethodInvoker)delegate
                            {
                                if (!dialog.IsDisposed)
                                {
                                    action();
                                }
                            });
                        }
                    }
                    catch
                    {
                    }
                };

                Action appendLog = delegate { };
                appendLog = () => { };
                Action<string> addLog = text =>
                {
                    ui(() =>
                    {
                        if (outputBox.TextLength > 0)
                        {
                            outputBox.AppendText(Environment.NewLine);
                        }

                        outputBox.AppendText(text ?? string.Empty);
                    });
                };

                Action saveDefaults = delegate
                {
                    _settings.OfficeCommunicationWindowTitle = string.IsNullOrWhiteSpace(titleBox.Text)
                        ? DefaultOfficeCommunicationWindowTitle
                        : titleBox.Text.Trim();
                    _settings.OfficeCommunicationWindowDescription = string.IsNullOrWhiteSpace(descriptionBox.Text)
                        ? DefaultOfficeCommunicationWindowDescription
                        : descriptionBox.Text.Trim();
                    _settings.OfficeCommunicationTestUrl = (directoryBox.Text ?? string.Empty).Trim();
                    _settings.OfficeCommunicationTestPayload = noteBox.Text ?? string.Empty;
                    _settings.Save();
                };

                Action<List<string>> rebuildSteps = executablePaths =>
                {
                    stepFlow.SuspendLayout();
                    stepFlow.Controls.Clear();
                    stepIcons.Clear();
                    stepTimes.Clear();

                    for (var i = 0; i < executablePaths.Count; i++)
                    {
                        var executablePath = executablePaths[i];
                        var card = new Panel
                        {
                            Width = 146,
                            Height = 66,
                            Margin = new Padding(8, 8, 0, 8),
                            BackColor = Color.White
                        };
                        StyleTaskFlowSurface(card);
                        var iconLabel = new Label
                        {
                            Text = "○",
                            Location = new Point(8, 16),
                            Size = new Size(28, 28),
                            TextAlign = ContentAlignment.MiddleCenter,
                            Font = new Font("Microsoft YaHei UI", 14F, FontStyle.Bold, GraphicsUnit.Point),
                            ForeColor = Color.FromArgb(100, 116, 139)
                        };
                        var nameLabel = new Label
                        {
                            Text = Path.GetFileNameWithoutExtension(executablePath),
                            Location = new Point(42, 9),
                            Size = new Size(94, 18),
                            TextAlign = ContentAlignment.MiddleLeft,
                            AutoEllipsis = true,
                            ForeColor = Color.FromArgb(15, 23, 42)
                        };
                        var timeLabel = new Label
                        {
                            Text = "--",
                            Location = new Point(42, 31),
                            Size = new Size(96, 22),
                            TextAlign = ContentAlignment.MiddleLeft,
                            ForeColor = Color.DimGray
                        };
                        timeLabel.Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point);
                        card.Controls.Add(iconLabel);
                        card.Controls.Add(nameLabel);
                        card.Controls.Add(timeLabel);
                        stepFlow.Controls.Add(card);
                        stepIcons.Add(iconLabel);
                        stepTimes.Add(timeLabel);

                        if (i < executablePaths.Count - 1)
                        {
                            stepFlow.Controls.Add(CreateTaskFlowConnector());
                        }
                    }

                    if (executablePaths.Count == 0)
                    {
                        var emptyLabel = new Label
                        {
                            Text = "当前目录没有找到可执行程序(.exe/.bat/.cmd/.ps1)",
                            AutoSize = false,
                            Width = 760,
                            Height = 60,
                            TextAlign = ContentAlignment.MiddleCenter,
                            ForeColor = Color.DimGray,
                            BackColor = Color.White
                        };
                        stepFlow.Controls.Add(emptyLabel);
                    }

                    stepFlow.ResumeLayout();
                };

                Action<int, string, Color, string, Color> setStepState = (index, symbol, symbolColor, elapsedText, elapsedColor) =>
                {
                    if (index < 0 || index >= stepIcons.Count || index >= stepTimes.Count)
                    {
                        return;
                    }

                    stepIcons[index].Text = symbol;
                    stepIcons[index].ForeColor = symbolColor;
                    stepTimes[index].Text = elapsedText;
                    stepTimes[index].ForeColor = elapsedColor;
                };

                Action updateElapsed = () =>
                {
                    totalElapsedLabel.Text = "总耗时：" + FormatCommunicationElapsed(totalStopwatch.Elapsed);
                };

                Action setRunningUi = () =>
                {
                    rerunButton.Enabled = false;
                    saveDefaultsButton.Enabled = false;
                    titleBox.Enabled = false;
                    descriptionBox.Enabled = false;
                    directoryBox.Enabled = false;
                    noteBox.Enabled = false;
                };

                Action setIdleUi = () =>
                {
                    rerunButton.Enabled = true;
                    saveDefaultsButton.Enabled = true;
                    titleBox.Enabled = true;
                    descriptionBox.Enabled = true;
                    directoryBox.Enabled = true;
                    noteBox.Enabled = true;
                };

                Action startExecution = delegate
                {
                    if (isRunning)
                    {
                        return;
                    }

                    var executablePaths = GetOfficeCommunicationExecutablePaths(directoryBox.Text);
                    rebuildSteps(executablePaths);
                    outputBox.Clear();
                    totalStopwatch.Reset();
                    totalElapsedLabel.Text = "总耗时：0.00s";

                    if (executablePaths.Count == 0)
                    {
                        summaryLabel.Text = "未找到可执行程序";
                        summaryLabel.ForeColor = Color.FromArgb(185, 28, 28);
                        addLog("未找到可执行程序。请检查“执行目录”是否正确，并确认目录内存在 .exe / .bat / .cmd / .ps1 文件。");
                        return;
                    }

                    cancellationRequested = false;
                    isRunning = true;
                    setRunningUi();
                    summaryLabel.Text = string.Format("准备执行，共 {0} 步", executablePaths.Count);
                    summaryLabel.ForeColor = Color.DimGray;
                    totalStopwatch.Start();

                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        var successCount = 0;
                        var failedCount = 0;

                        for (var i = 0; i < executablePaths.Count; i++)
                        {
                            if (cancellationRequested)
                            {
                                break;
                            }

                            var stepIndex = i;
                            var stepDisplayIndex = i + 1;
                            var executablePath = executablePaths[stepIndex];
                            addLog("开始执行：" + executablePath);
                            ui(() =>
                            {
                                summaryLabel.Text = string.Format("正在执行 {0}/{1}", stepDisplayIndex, executablePaths.Count);
                                summaryLabel.ForeColor = Color.FromArgb(30, 64, 175);
                                setStepState(stepIndex, "◔", Color.FromArgb(37, 99, 235), "执行中", Color.FromArgb(37, 99, 235));
                                updateElapsed();
                            });

                            var stepWatch = Stopwatch.StartNew();
                            var success = false;
                            var resultText = string.Empty;

                            try
                            {
                                using (var process = new Process())
                                {
                                    process.StartInfo = BuildOfficeCommunicationProcessStartInfo(executablePath);
                                    process.Start();
                                    currentProcess = process;
                                    currentProcessId = process.Id;

                                    while (!process.WaitForExit(120))
                                    {
                                        if (cancellationRequested)
                                        {
                                            TerminateProcessTree(currentProcessId);
                                            resultText = "已取消";
                                            break;
                                        }

                                        ui(updateElapsed);
                                    }

                                    if (!cancellationRequested)
                                    {
                                        if (!WaitForProcessTreeExit(currentProcessId, () => cancellationRequested, () => ui(updateElapsed)))
                                        {
                                            resultText = "已取消";
                                        }
                                        else
                                        {
                                            success = process.ExitCode == 0;
                                            resultText = success
                                                ? "退出码 0，任务树完成"
                                                : "退出码 " + process.ExitCode + "，任务树结束";
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                resultText = ex.Message;
                            }
                            finally
                            {
                                currentProcess = null;
                                currentProcessId = 0;
                                stepWatch.Stop();
                            }

                            if (cancellationRequested)
                            {
                                failedCount++;
                                ui(() => setStepState(stepIndex, "⊗", Color.FromArgb(185, 28, 28), "已取消", Color.FromArgb(185, 28, 28)));
                                addLog("已取消：" + executablePath);
                                break;
                            }

                            var elapsedText = FormatCommunicationElapsed(stepWatch.Elapsed);
                            if (success)
                            {
                                successCount++;
                                ui(() => setStepState(stepIndex, "☑", Color.FromArgb(21, 128, 61), elapsedText, Color.FromArgb(21, 128, 61)));
                                addLog(string.Format("成功：{0} | 耗时 {1} | {2}", executablePath, elapsedText, resultText));
                            }
                            else
                            {
                                failedCount++;
                                ui(() => setStepState(stepIndex, "⊗", Color.FromArgb(185, 28, 28), elapsedText, Color.FromArgb(185, 28, 28)));
                                addLog(string.Format("失败：{0} | 耗时 {1} | {2}", executablePath, elapsedText, resultText));
                            }

                            ui(updateElapsed);

                            if (i < executablePaths.Count - 1 && _settings.OfficeCommunicationStepDelayMs > 0)
                            {
                                var waited = 0;
                                while (waited < _settings.OfficeCommunicationStepDelayMs && !cancellationRequested)
                                {
                                    Thread.Sleep(50);
                                    waited += 50;
                                    ui(updateElapsed);
                                }
                            }
                        }

                        totalStopwatch.Stop();
                        ui(() =>
                        {
                            updateElapsed();
                            if (cancellationRequested)
                            {
                                summaryLabel.Text = "执行已取消";
                                summaryLabel.ForeColor = Color.FromArgb(185, 28, 28);
                            }
                            else
                            {
                                summaryLabel.Text = string.Format("执行完成：成功 {0}，失败 {1}", successCount, failedCount);
                                summaryLabel.ForeColor = failedCount == 0
                                    ? Color.FromArgb(21, 128, 61)
                                    : Color.FromArgb(185, 28, 28);
                            }

                            isRunning = false;
                            setIdleUi();
                        });
                    });
                };

                rerunButton.Click += (_, __) => startExecution();
                saveDefaultsButton.Click += (_, __) => saveDefaults();
                copyAllButton.Click += (_, __) =>
                {
                    var allText = string.Format(
                        "窗口标题：{0}\r\n\r\n窗口说明：\r\n{1}\r\n\r\n执行目录：\r\n{2}\r\n\r\n附加文字/备注：\r\n{3}\r\n\r\n状态：{4}\r\n{5}\r\n\r\n执行输出：\r\n{6}",
                        string.IsNullOrWhiteSpace(titleBox.Text) ? DefaultOfficeCommunicationWindowTitle : titleBox.Text.Trim(),
                        descriptionBox.Text ?? string.Empty,
                        directoryBox.Text ?? string.Empty,
                        noteBox.Text ?? string.Empty,
                        summaryLabel.Text ?? string.Empty,
                        totalElapsedLabel.Text ?? string.Empty,
                        outputBox.Text ?? string.Empty);
                    try
                    {
                        CopyTextToClipboard(allText);
                    }
                    catch (Exception ex)
                    {
                        ShowErrorMessage("复制任务流内容失败：" + ex.Message, "任务流窗口");
                    }
                };
                closeButton.Click += (_, __) => dialog.Close();
                escTipLink.Click += (_, __) => dialog.Close();

                dialog.FormClosing += (_, __) =>
                {
                    cancellationRequested = true;
                    try
                    {
                        if (currentProcessId > 0)
                        {
                            TerminateProcessTree(currentProcessId);
                        }
                        else if (currentProcess != null && !currentProcess.HasExited)
                        {
                            currentProcess.Kill();
                        }
                    }
                    catch
                    {
                    }

                    saveDefaults();
                };

                dialog.KeyDown += (_, e) =>
                {
                    if (e != null && e.KeyCode == Keys.Escape)
                    {
                        e.SuppressKeyPress = true;
                        e.Handled = true;
                        dialog.Close();
                    }
                };

                topCard.Controls.AddRange(new Control[]
                {
                    titleBox, escTipLink, totalElapsedLabel, descriptionLabel, descriptionBox,
                    directoryLabel, directoryBox, noteLabel, noteBox
                });
                stepCard.Controls.AddRange(new Control[]
                {
                    summaryLabel, stepFlow
                });
                outputCard.Controls.AddRange(new Control[]
                {
                    outputLabel, outputBox
                });

                dialog.Controls.AddRange(new Control[]
                {
                    topCard, stepCard, outputCard, rerunButton, saveDefaultsButton, copyAllButton, closeButton
                });
                ApplyDialogDpiScaling(dialog, new Size(860, 674));

                dialog.Shown += (_, __) => startExecution();
                dialog.ShowDialog(this);
            }
        }
    }
}
