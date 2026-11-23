using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32;

// ---------------------------------------------------------------------------
// OfficeAI 唤醒器 (Office Waker)
//
// 这是一个**极小的 COM 加载项**，唯一职责是：Word / Excel / PowerPoint 启动时
// 检查本地网关是否已运行，如果没有就把它拉起来。它不提供任何 UI，侧边栏仍
// 然由现有的 Office Web Add-in（由网关提供页面）负责。
//
// 生命周期由此与办公软件严格绑定：
//   打开 Office → 唤醒器 OnConnection → 拉起网关
//   关闭 Office → 网关的 LifecycleMonitor 检测到进程全部消失 → 自动退出
//
// 注册（per-user，免管理员）：
//   HKCU\Software\Microsoft\Office\<App>\Addins\OfficeAI.Waker
//       LoadBehavior = 3   (启动时加载)
// ---------------------------------------------------------------------------

namespace OfficeAI
{
    [ComVisible(true)]
    [Guid("4F1D9C3A-8B27-4E6D-9C41-0A7B5E2D8F31")]
    [ProgId("OfficeAI.Waker")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Waker : IDTExtensibility2
    {
        private const string GatewayUrl = "http://127.0.0.1:18080";
        private const string LogName = "waker.log";

        private static readonly object Gate = new object();
        private static bool _launched;

        /// <summary>
        /// 关键：OnConnection 里**什么都不做**，直接返回。
        /// 此时 Office 还在启动中（splash 阶段），任何 I/O 都会计入加载项耗时，
        /// 累计超过阈值就会被 Office 判定为“慢加载项”自动禁用。
        /// 真正的探活/启动放到 OnStartupComplete（界面已就绪）之后。
        /// </summary>
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            // intentionally empty
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom)
        {
            StartWatchdog();
        }
        public void OnBeginShutdown(ref Array custom) { }

        /// <summary>
        /// 在后台线程上探活并按需启动：Office 主线程立刻返回，绝不等待网络。
        /// 1.5 秒延迟让 Office 先把界面画完；随后每 5 秒补探一次，
        /// 覆盖“网关刚被关掉、用户继续用 Office”的场景（最多 12 次，约 1 分钟）。
        /// </summary>
        private static void StartWatchdog()
        {
            var thread = new Thread(() =>
            {
                try
                {
                    Thread.Sleep(1500);
                    for (int attempt = 0; attempt < 12; attempt++)
                    {
                        if (IsGatewayAlive())
                        {
                            return;
                        }
                        LaunchIfNeeded();
                        Thread.Sleep(5000);
                    }
                }
                catch (Exception ex)
                {
                    Log("watchdog failed: " + ex.Message);
                }
            });

            thread.IsBackground = true; // 不阻止 Office 退出
            thread.Name = "OfficeAI.Waker.Watchdog";
            thread.Start();
        }

        /// <summary>
        /// 探活 → 需要时拉起 Start-OfficeAI.bat（由后台线程调用，阻塞无所谓）。
        /// </summary>
        private static void LaunchIfNeeded()
        {
            if (IsGatewayAlive())
            {
                Log("gateway already running");
                return;
            }

            lock (Gate)
            {
                if (_launched) return; // 同一进程内只尝试一次，避免拖慢启动
                _launched = true;
            }

            string launcher = ResolveLauncher();
            if (launcher == null)
            {
                Log("no launcher found");
                return;
            }

            try
            {
                if (launcher.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
                {
                    // 用 cmd /c start /min 启动，避免弹出黑窗口
                    var psi = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = "/c start \"OfficeAI\" /min \"" + launcher + "\" --launched-by office",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        WorkingDirectory = Path.GetDirectoryName(launcher) ?? Environment.CurrentDirectory
                    };
                    Process.Start(psi);
                }
                else
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = launcher,
                        Arguments = "--launched-by office",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    Process.Start(psi);
                }
                Log("launched: " + launcher);
            }
            catch (Exception ex)
            {
                Log("launch failed: " + ex.Message);
            }
        }

        private static bool IsGatewayAlive()
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(GatewayUrl + "/api/health");
                request.Timeout = 300;   // 300ms：服务未起来时立即失败，不拖慢 Office
                request.ReadWriteTimeout = 300;
                request.Method = "GET";
                request.Proxy = null;
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

        /// <summary>
        /// 查找启动程序，顺序与网关的 resolveLaunchTarget 保持一致：
        /// 注册表配置 → 环境变量 → OfficeAI.exe 常见安装位置 → Start-OfficeAI.bat
        /// </summary>
        private static string ResolveLauncher()
        {
            string fromRegistry = ReadRegistryConfig("gatewayLaunchCommand");
            if (!string.IsNullOrEmpty(fromRegistry) && File.Exists(fromRegistry))
            {
                return fromRegistry;
            }

            string fromEnv = Environment.GetEnvironmentVariable("OFFICEAI_LAUNCHER");
            if (!string.IsNullOrEmpty(fromEnv) && File.Exists(fromEnv))
            {
                return fromEnv;
            }

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
                    if (!string.IsNullOrEmpty(candidate) && File.Exists(candidate))
                    {
                        return Path.GetFullPath(candidate);
                    }
                }
                catch
                {
                    // 忽略非法路径
                }
            }

            return null;
        }

        /// <summary>读取网关写入的 config.json 中的 gatewayLaunchCommand</summary>
        private static string ReadRegistryConfig(string key)
        {
            try
            {
                string configPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "OfficeAI",
                    "config.json");
                if (!File.Exists(configPath))
                {
                    return null;
                }

                string json = File.ReadAllText(configPath, Encoding.UTF8);
                int idx = json.IndexOf("\"" + key + "\"", StringComparison.Ordinal);
                if (idx < 0) return null;

                int colon = json.IndexOf(':', idx);
                if (colon < 0) return null;

                int firstQuote = json.IndexOf('"', colon + 1);
                if (firstQuote < 0) return null;

                var value = new StringBuilder();
                for (int i = firstQuote + 1; i < json.Length; i++)
                {
                    char c = json[i];
                    if (c == '\\' && i + 1 < json.Length)
                    {
                        value.Append(json[++i]); // 解转义（\\\\ → \\）
                        continue;
                    }
                    if (c == '"') break;
                    value.Append(c);
                }

                string result = value.ToString().Trim();
                return result.Length > 0 ? result : null;
            }
            catch
            {
                return null;
            }
        }

        private static void Log(string message)
        {
            try
            {
                // 简单滚动，避免日志无限增长
                string logPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "OfficeAI",
                    LogName);
                if (System.IO.File.Exists(logPath) && new System.IO.FileInfo(logPath).Length > 128 * 1024)
                {
                    System.IO.File.Delete(logPath);
                }
            }
            catch
            {
                // ignore
            }

            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "OfficeAI");
                Directory.CreateDirectory(dir);
                File.AppendAllText(
                    Path.Combine(dir, LogName),
                    DateTime.Now.ToString("s") + "  " + message + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch
            {
                // 日志失败不能影响 Office 启动
            }
        }
    }

    // -----------------------------------------------------------------------
    // IDTExtensibility2：从 Extensibility.dll 手写声明，避免额外引用/构建步骤
    // -----------------------------------------------------------------------
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType(TypeLibTypeFlags.FDispatchable)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            [In] ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] ext_DisconnectMode RemoveMode,
            [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1
    }
}
