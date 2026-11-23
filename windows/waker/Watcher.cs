using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

// ---------------------------------------------------------------------------
// OfficeAI 后台监听器 (Office Watcher)
//
// 与 Office 唤醒器（COM 加载项）的区别：
//   * 唤醒器：Word/Excel/PPT 启动时由 Office 加载，**在被宿主进程内**执行探活。
//   * 监听器：本程序是一个独立的极轻量常驻进程，**完全不进 Office 进程**，
//     只轮询进程列表；一旦发现 Office/WPS 启动且网关未运行，就把网关拉起。
//
// 这样 Office 的启动路径上没有任何第三方代码，绝对不会被"慢加载项"机制禁用；
// 代价是需要一个常驻进程（约 10MB）和一条登录自启的计划任务。
//
// 资源占用：每 2 秒一次 Process.GetProcesses() 过滤（约 2-4ms CPU），空闲时
// 内存稳定在 10MB 左右，无 UI、无托盘图标。
// ---------------------------------------------------------------------------

namespace OfficeAI
{
    internal static class Watcher
    {
        private const string GatewayUrl = "http://127.0.0.1:18080";
        private const int PollMs = 2000;
        private const int ProbeTimeoutMs = 400;

        private static readonly string[] OfficeProcesses =
        {
            "winword", "excel", "powerpnt", "wps", "et", "wpp", "wpsoffice"
        };

        private static string _dataDir;
        private static string _pidFile;
        private static string _lastLaunchTarget;
        private static DateTime _lastLaunchAt = DateTime.MinValue;

        private static int Main(string[] args)
        {
            _dataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OfficeAI");
            Directory.CreateDirectory(_dataDir);
            _pidFile = Path.Combine(_dataDir, "watcher.pid");

            if (args.Length > 0 && args[0] == "--once")
            {
                // 手动跑一次：便于排查"为什么没自动启动"
                bool running = HasOfficeProcess(out List<string> names);
                Log("--once: office=" + (running ? string.Join(",", names) : "none")
                    + " gateway=" + (IsGatewayAlive() ? "up" : "down"));
                if (running && !IsGatewayAlive() && ReadLaunchTrigger() != "manual")
                {
                    Log("--once: launch=" + (LaunchGateway() ?? "failed"));
                }
                return 0;
            }

            // 单实例：避免计划任务重复拉起
            using (var mutex = new Mutex(true, "Global\\OfficeAI.Watcher", out bool isFirst))
            {
                if (!isFirst)
                {
                    Log("another watcher instance is already running, exiting");
                    return 0;
                }

                try
                {
                    File.WriteAllText(_pidFile, Process.GetCurrentProcess().Id.ToString(), Encoding.UTF8);
                }
                catch
                {
                    // 忽略
                }

                Log("watcher started, pid=" + Process.GetCurrentProcess().Id);
                AppDomain.CurrentDomain.ProcessExit += (s, e) => Cleanup();
                AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                {
                    Log("unhandled: " + e.ExceptionObject);
                };

                Run();
                Cleanup();
                return 0;
            }
        }

        private static void Run()
        {
            bool officeWasRunning = false;
            DateTime lastHeartbeat = DateTime.MinValue;

            while (true)
            {
                try
                {
                    // 用户在配置中心切到 manual / always 时，监听器应停止介入
                    string trigger = ReadLaunchTrigger();
                    if (trigger == "off")
                    {
                        Log("launchTrigger=off, watcher exits");
                        return;
                    }

                    bool officeRunning = HasOfficeProcess(out List<string> names);

                    if (officeRunning && !officeWasRunning)
                    {
                        Log("office started: " + string.Join(",", names));
                    }
                    officeWasRunning = officeRunning;

                    if (officeRunning && !IsGatewayAlive())
                    {
                        // 冷却时间：拉不起来时不要每 2 秒重试轰炸
                        if ((DateTime.UtcNow - _lastLaunchAt).TotalSeconds > 30)
                        {
                            _lastLaunchAt = DateTime.UtcNow;
                            string launched = LaunchGateway();
                            Log("gateway launch -> " + (launched ?? "failed (no target)"));
                        }
                    }

                    if ((DateTime.UtcNow - lastHeartbeat).TotalSeconds > 30)
                    {
                        lastHeartbeat = DateTime.UtcNow;
                        TouchPidFile();
                    }
                }
                catch (Exception ex)
                {
                    Log("loop error: " + ex.Message);
                }

                Thread.Sleep(PollMs);
            }
        }

        private static void Cleanup()
        {
            try
            {
                if (File.Exists(_pidFile)) File.Delete(_pidFile);
            }
            catch
            {
                // 忽略
            }
        }

        private static void TouchPidFile()
        {
            try
            {
                File.WriteAllText(_pidFile, Process.GetCurrentProcess().Id.ToString() + "\n" + DateTime.UtcNow.ToString("o"), Encoding.UTF8);
            }
            catch
            {
                // 忽略
            }
        }

        /// <summary>只统计主程序，忽略 wpscloudsvr / wpsupdate 等辅助进程</summary>
        private static bool HasOfficeProcess(out List<string> matched)
        {
            matched = new List<string>();
            var wanted = new HashSet<string>(OfficeProcesses, StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (Process process in Process.GetProcesses())
                {
                    try
                    {
                        string name = process.ProcessName;
                        if (wanted.Contains(name))
                        {
                            matched.Add(name.ToLowerInvariant());
                        }
                    }
                    catch
                    {
                        // 进程可能已退出
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch
            {
                return false;
            }
            return matched.Count > 0;
        }

        private static bool IsGatewayAlive()
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(GatewayUrl + "/api/health");
                request.Timeout = ProbeTimeoutMs;
                request.ReadWriteTimeout = ProbeTimeoutMs;
                request.Method = "GET";
                request.Proxy = null;
                request.KeepAlive = false;
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    return response.StatusCode == HttpStatusCode.OK;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>读取 config.json 的 launchMode / gatewayLaunchCommand</summary>
        private static string ReadJsonField(string path, string key)
        {
            try
            {
                if (!File.Exists(path)) return null;
                string json = File.ReadAllText(path, Encoding.UTF8);
                int idx = json.IndexOf("\"" + key + "\"", StringComparison.Ordinal);
                if (idx < 0) return null;
                int colon = json.IndexOf(':', idx);
                if (colon < 0) return null;
                int firstQuote = json.IndexOf('"', colon + 1);
                if (firstQuote < 0)
                {
                    // 布尔/数字：读到最近的 , 或 }
                    var raw = new StringBuilder();
                    for (int i = colon + 1; i < json.Length; i++)
                    {
                        char c = json[i];
                        if (c == ',' || c == '}' || c == '\n') break;
                        raw.Append(c);
                    }
                    return raw.ToString().Trim();
                }
                var value = new StringBuilder();
                for (int i = firstQuote + 1; i < json.Length; i++)
                {
                    char c = json[i];
                    if (c == '\\' && i + 1 < json.Length) { value.Append(json[++i]); continue; }
                    if (c == '"') break;
                    value.Append(c);
                }
                return value.ToString().Trim();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>manual 模式不自动启动；always 模式网关本来就常驻</summary>
        private static string ReadLaunchTrigger()
        {
            string configPath = Path.Combine(_dataDir, "config.json");
            string mode = ReadJsonField(configPath, "launchMode");
            if (string.Equals(mode, "manual", StringComparison.OrdinalIgnoreCase)) return "manual";
            string trigger = ReadJsonField(configPath, "launchTrigger");
            if (string.Equals(trigger, "waker", StringComparison.OrdinalIgnoreCase)) return "off";
            return "watcher";
        }

        private static string ResolveLauncher()
        {
            string configPath = Path.Combine(_dataDir, "config.json");
            string configured = ReadJsonField(configPath, "gatewayLaunchCommand");
            if (!string.IsNullOrEmpty(configured) && File.Exists(configured)) return configured;

            string env = Environment.GetEnvironmentVariable("OFFICEAI_LAUNCHER");
            if (!string.IsNullOrEmpty(env) && File.Exists(env)) return env;

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string[] candidates =
            {
                Path.Combine(localAppData, "Programs", "OfficeAI", "OfficeAI.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OfficeAI", "OfficeAI.exe"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OfficeAI.exe"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Start-OfficeAI.bat"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "Start-OfficeAI.bat")
            };
            foreach (string candidate in candidates)
            {
                try
                {
                    if (!string.IsNullOrEmpty(candidate) && File.Exists(candidate)) return Path.GetFullPath(candidate);
                }
                catch
                {
                    // 忽略非法路径
                }
            }
            return null;
        }

        private static string LaunchGateway()
        {
            string launcher = ResolveLauncher();
            if (string.IsNullOrEmpty(launcher)) return null;
            _lastLaunchTarget = launcher;

            try
            {
                if (launcher.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = "/c start \"OfficeAI\" /min \"" + launcher + "\" --launched-by watcher",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        WorkingDirectory = Path.GetDirectoryName(launcher) ?? Environment.CurrentDirectory
                    });
                }
                else
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = launcher,
                        Arguments = "--launched-by watcher",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    });
                }
                return launcher;
            }
            catch (Exception ex)
            {
                Log("launch failed: " + ex.Message);
                return null;
            }
        }

        private static void Log(string message)
        {
            try
            {
                string logPath = Path.Combine(_dataDir, "watcher.log");
                // 简单滚动，避免无限增长
                if (File.Exists(logPath) && new FileInfo(logPath).Length > 256 * 1024)
                {
                    File.Copy(logPath, logPath + ".1", true);
                    File.Delete(logPath);
                }
                File.AppendAllText(
                    logPath,
                    DateTime.Now.ToString("s") + "  " + message + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch
            {
                // 日志失败不影响主循环
            }
        }
    }
}
