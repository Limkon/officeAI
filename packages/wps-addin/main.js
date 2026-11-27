/* global wps, et, wpp, Application */
/**
 * OfficeAI WPS 加载项（本地离线版）
 *
 * 该文件被安装到 %APPDATA%\kingsoft\wps\jsaddons\officeai_<version>\ 下，
 * 属于「离线加载项」，WPS 启动时直接用 file:// 读取，不依赖本地网关。
 *
 * 因此它可以承担「按需启动」的职责：
 *   1. 点击功能区的任意按钮时，先探活 http://127.0.0.1:18080/api/health
 *   2. 服务未启动 → 用 wps.OAAssist.ShellExecute 拉起 OfficeAI（只弹一次确认框）
 *   3. 服务已启动 → 直接打开侧边栏 / 配置中心
 *
 * 安装时（Install-OfficeAI.ps1）会把下面的 __OFFICEAI_*__ 占位符替换为真实值。
 */
(function () {
  var GATEWAY_URL = '__OFFICEAI_GATEWAY_URL__';
  var LAUNCH_COMMAND = '__OFFICEAI_LAUNCH_COMMAND__';
  var LAUNCH_URL = '__OFFICEAI_LAUNCH_URL__';
  var BOOTSTRAP_URL = '__OFFICEAI_BOOTSTRAP_URL__';

  if (GATEWAY_URL.indexOf('__OFFICEAI_') === 0) {
    // Running from the repo / gateway served copy: fall back to defaults.
    GATEWAY_URL = 'http://127.0.0.1:18080';
    LAUNCH_COMMAND = '';
    LAUNCH_URL = 'officeai://activate/' + encodeURIComponent(GATEWAY_URL + '/start?app=wps');
    BOOTSTRAP_URL = GATEWAY_URL + '/wps/bootstrap.html';
  }

  var TASKPANE_KEY = 'officeai_taskpane_id';
  var HEALTH_TIMEOUT_MS = 1500;
  var LAUNCH_POLL_MS = 1200;
  var LAUNCH_POLL_ATTEMPTS = 30;
  var launching = false;

  // ---------------------------------------------------------------- host API
  function getApp() {
    if (typeof window.Application !== 'undefined' && window.Application) {
      return window.Application;
    }
    if (typeof wps !== 'undefined' && wps.WpsApplication) {
      return wps.WpsApplication();
    }
    if (typeof et !== 'undefined' && et.EtApplication) {
      return et.EtApplication();
    }
    if (typeof wpp !== 'undefined' && wpp.WppApplication) {
      return wpp.WppApplication();
    }
    return null;
  }

  function detectWpsHostType() {
    var app = getApp();
    if (!app) return 'wps-writer';
    try {
      var name = String(app.Name || '').toLowerCase();
      if (name.indexOf('et') !== -1 || name.indexOf('spread') !== -1 || app.ActiveWorkbook) {
        return 'wps-spreadsheets';
      }
      if (name.indexOf('wpp') !== -1 || name.indexOf('present') !== -1 || app.ActivePresentation) {
        return 'wps-presentation';
      }
    } catch (e) {
      // ignore
    }
    return 'wps-writer';
  }

  // ------------------------------------------------------------ health check
  function gatewayAlive() {
    return new Promise(function (resolve) {
      var settled = false;
      var timer = null;
      function finish(ok) {
        if (settled) return;
        settled = true;
        if (timer) clearTimeout(timer);
        resolve(ok);
      }
      timer = setTimeout(function () {
        finish(false);
      }, HEALTH_TIMEOUT_MS);

      try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', GATEWAY_URL + '/api/health?_t=' + Date.now(), true);
        xhr.onreadystatechange = function () {
          if (xhr.readyState !== 4) return;
          finish(xhr.status >= 200 && xhr.status < 300);
        };
        xhr.onerror = function () {
          finish(false);
        };
        xhr.send();
      } catch (e) {
        finish(false);
      }
    });
  }

  function sleep(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  function showLaunchFailed(hostType) {
    var app = getApp();
    var cmd = LAUNCH_COMMAND || LAUNCH_URL || 'OfficeAI.exe';
    var msg =
      'OfficeAI 本地服务启动失败。\n\n' +
      '可能原因：\n' +
      '1. WPS 安全设置拦截了外部程序启动（首次会弹一次确认，请点“允许”）\n' +
      '2. 安装路径不存在或端口 18080 被占用\n\n' +
      '请尝试：\n' +
      '• 手动启动：按 Win 键搜索“OfficeAI”运行，然后重试按钮\n' +
      '• 重启 WPS 后再点一次\n' +
      '• 检查 Office 面板 → 配置中心 → 运行状态\n\n' +
      '启动命令：' + cmd + '\n网关地址：' + GATEWAY_URL;
    try {
      if (app && typeof app.Alert === 'function') {
        app.Alert(msg);
        return;
      }
      if (typeof wps !== 'undefined' && typeof wps.Alert === 'function') {
        wps.Alert(msg);
        return;
      }
    } catch (e) {}
    try {
      alert(msg);
    } catch (e2) {}
    // 兜底：仍尝试用 ShowDialog 打开引导页，方便用户一键重试（即使网关未启动，也能看到本地 launcher 的重试逻辑）
    try {
      if (app && typeof app.ShowDialog === 'function') {
        var fallbackUrl = BOOTSTRAP_URL + '?app=' + encodeURIComponent(hostType || 'wps-writer');
        app.ShowDialog(fallbackUrl, 'OfficeAI 启动失败', 760, 520, false);
      }
    } catch (e3) {}
  }

  /**
   * Starts the OfficeAI client. WPS shows a one-time security prompt the first
   * time an add-in executes an external program, so we open the gateway's
   * friendly launcher page first: the user sees *why* OfficeAI wants to start.
   * If the user has already launched it (or the confirm dialog is suppressed),
   * the health check simply succeeds.
   */
  function launchGateway(actionId, hostType) {
    var app = getApp();
    var targetUrl =
      BOOTSTRAP_URL + '?app=' + encodeURIComponent(hostType) + (actionId ? '&action=' + encodeURIComponent(actionId) : '');
    var cmd = LAUNCH_COMMAND || LAUNCH_URL;
    var started = false;
    var lastErr = '';

    if (!cmd) {
      // 未能解析到启动命令时，直接提示用户手动启动，避免静默无响应
      showLaunchFailed(hostType);
      return false;
    }

    try {
      if (typeof wps !== 'undefined' && wps.OAAssist && typeof wps.OAAssist.ShellExecute === 'function') {
        wps.OAAssist.ShellExecute(cmd);
        started = true;
      }
    } catch (e) {
      lastErr = String(e && e.message ? e.message : e);
      started = false;
    }

    if (!started && app && app.OAAssist) {
      try {
        app.OAAssist.ShellExecute(cmd);
        started = true;
      } catch (e) {
        lastErr = String(e && e.message ? e.message : e);
        started = false;
      }
    }

    // 即使 ShellExecute 返回成功，也给用户一个可视的“正在启动”反馈，避免“点了没反应”错觉
    setTimeout(function () {
      try {
        if (app && typeof app.ShowDialog === 'function' && GATEWAY_URL.indexOf('http') === 0) {
          // 只有网关地址是 http 时才尝试打开引导页；file:// 场景下由 waitForGateway 后的逻辑负责
          app.ShowDialog(
            targetUrl,
            started ? 'OfficeAI 正在启动' : 'OfficeAI 启动中',
            760,
            560,
            false
          );
        } else if (!started) {
          showLaunchFailed(hostType);
        }
      } catch (e) {
        if (!started) showLaunchFailed(hostType);
      }
    }, started ? 400 : 0);

    if (!started && lastErr) {
      // 记录到控制台，方便在 WPS 开发者工具中排查
      try { console.log('[OfficeAI] ShellExecute 失败:', lastErr, 'cmd:', cmd); } catch {}
    }

    return started;
  }

  function waitForGateway() {
    var attempt = 0;
    return new Promise(function (resolve) {
      function step() {
        gatewayAlive().then(function (ok) {
          if (ok) return resolve(true);
          attempt += 1;
          if (attempt >= LAUNCH_POLL_ATTEMPTS) return resolve(false);
          setTimeout(step, LAUNCH_POLL_MS);
        });
      }
      step();
    });
  }

  // --------------------------------------------------------------- taskpane
  function openOrUpdateTaskPane(actionId) {
    var app = getApp();
    if (!app) return;

    var hostType = detectWpsHostType();
    var targetUrl =
      GATEWAY_URL +
      '/?host=wps&app=' +
      encodeURIComponent(hostType) +
      (actionId ? '&action=' + encodeURIComponent(actionId) : '') +
      '&t=' +
      Date.now();

    // Try reusing existing TaskPane in its own try/catch so stale IDs fall back to CreateTaskPane
    try {
      var tsId = null;
      if (app.PluginStorage && typeof app.PluginStorage.getItem === 'function') {
        tsId = app.PluginStorage.getItem(TASKPANE_KEY);
      }
      if (tsId && typeof app.GetTaskPane === 'function') {
        var existingPane = app.GetTaskPane(tsId);
        if (existingPane) {
          existingPane.Visible = true;
          if (actionId && existingPane.Navigate) {
            existingPane.Navigate(targetUrl);
          }
          return;
        }
      }
    } catch (e) {
      // ignore stale taskpane reference and recreate below
    }

    try {
      if (typeof app.CreateTaskPane === 'function') {
        var pane = app.CreateTaskPane(targetUrl, 'OfficeAI 智能助手');
        if (pane) {
          pane.DockPosition = 2; // msoCTPDockPositionRight
          pane.Width = 400;
          pane.Visible = true;
          if (app.PluginStorage && typeof app.PluginStorage.setItem === 'function') {
            app.PluginStorage.setItem(TASKPANE_KEY, pane.ID);
          }
          return;
        }
      }
    } catch (err) {
      // fallback to dialog below
    }

    if (typeof app.ShowDialog === 'function') {
      app.ShowDialog(targetUrl, 'OfficeAI 智能助手', 960, 680, false);
    }
  }

  function openDashboard() {
    var app = getApp();
    var dashUrl = GATEWAY_URL + '/dashboard';
    if (app && typeof app.ShowDialog === 'function') {
      app.ShowDialog(dashUrl, 'OfficeAI 模型与集成配置中心', 1100, 760, false);
    } else {
      openOrUpdateTaskPane('');
    }
  }

  // --------------------------------------------------------------- lifecycle
  function ensureGatewayThen(actionId) {
    var hostType = detectWpsHostType();
    return gatewayAlive().then(function (ok) {
      if (ok) {
        try { openOrUpdateTaskPane(actionId); } catch (e) { try{ console.log('[OfficeAI] open pane failed', e);}catch{} }
        return true;
      }
      if (launching) {
        // 正在启动中，避免重复 ShellExecute，但也要给用户反馈
        try {
          var appTmp = getApp();
          if (appTmp && appTmp.Alert) appTmp.Alert('OfficeAI 正在启动中，请稍等 2-3 秒后重试…');
        } catch {}
        return false;
      }
      launching = true;
      var shellOk = launchGateway(actionId, hostType);
      if (!shellOk) {
        launching = false;
        showLaunchFailed(hostType);
        return false;
      }
      return waitForGateway().then(function (started) {
        launching = false;
        if (started) {
          try { openOrUpdateTaskPane(actionId); } catch (e) {}
          return true;
        }
        showLaunchFailed(hostType);
        return false;
      });
    }).catch(function (err) {
      launching = false;
      try { console.log('[OfficeAI] ensureGatewayThen 异常', err); } catch {}
      showLaunchFailed(detectWpsHostType());
      return false;
    });
  }

  // ------------------------------------------------------------ ribbon hooks
  window.OnAddinLoad = function (ribbonUI) {
    window.OfficeAIRibbonUI = ribbonUI;
    // 只做一次静默探活，不启动任何进程，避免打扰用户
    gatewayAlive();
    return true;
  };

  window.OnGetEnabled = function () {
    return true;
  };

  window.GetImage = function () {
    return GATEWAY_URL + '/wps/assets/icon-32.png';
  };

  window.OnAction = function (control) {
    try {
      var id = control && control.Id ? control.Id : 'btnOpenTaskPane';
      if (id === 'btnOpenDashboard') {
        ensureGatewayThen('').then(function (ok) {
          if (ok) {
            try { openDashboard(); } catch (e) { showLaunchFailed(detectWpsHostType()); }
          }
          // 失败时 ensureGatewayThen 内部已弹框，此处不再静默
        }).catch(function(){ showLaunchFailed(detectWpsHostType()); });
        return;
      }

      var actionMap = {
        btnOpenTaskPane: '',
        btnPolish: 'polish',
        btnTranslate: 'translate-en',
        btnSummarize: 'summarize',
        btnFormula: 'excel-formula',
        btnPptOutline: 'ppt-outline'
      };

      ensureGatewayThen(actionMap[id] || '').catch(function(){ showLaunchFailed(detectWpsHostType()); });
    } catch (err) {
      try { console.log('[OfficeAI] OnAction 异常', err); } catch {}
      showLaunchFailed(detectWpsHostType());
    }
  };

  /**
   * Used by index.html (the add-in entry page) and bootstrap.html: if the
   * gateway is reachable jump straight into the UI, otherwise render the
   * launcher so the user can start it with one click.
   */
  window.OfficeAIBootstrap = function (hostParam) {
    var host = hostParam === 'wps' ? detectWpsHostType() : hostParam || 'standalone';
    gatewayAlive().then(function (ok) {
      if (ok) {
        window.location.replace(GATEWAY_URL + '/?host=' + encodeURIComponent(host));
        return;
      }
      var shell = document.body || document.documentElement;
      shell.setAttribute(
        'style',
        'margin:0;font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;background:#f8fafc;' +
          'color:#0f172a;display:flex;align-items:center;justify-content:center;min-height:100vh;'
      );
      shell.innerHTML =
        '<div style="max-width:460px;padding:24px">' +
        '<div style="background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:20px">' +
        '<h1 style="font-size:17px;margin:0 0 8px">OfficeAI 服务未启动</h1>' +
        '<p style="font-size:13px;line-height:1.7;color:#475569;margin:0 0 14px">' +
        '点击下方按钮会在需要时自动拉起 OfficeAI 本地服务（不会开机自启），随后自动进入助手。' +
        '</p>' +
        '<button id="officeai-start" style="width:100%;border:0;border-radius:10px;padding:11px 14px;' +
        'font-size:14px;font-weight:600;color:#fff;background:#2563eb;cursor:pointer">' +
        '一键启动 OfficeAI</button>' +
        '<div id="officeai-state" style="font-size:12px;color:#64748b;margin-top:12px"></div>' +
        '</div></div>';

      var stateEl = document.getElementById('officeai-state');
      function setState(text) {
        if (stateEl) stateEl.textContent = text;
      }
      var btn = document.getElementById('officeai-start');
      if (btn) {
        btn.onclick = function () {
          setState('正在唤起 OfficeAI…');
          launchGateway('', host);
          waitForGateway().then(function (started) {
            if (started) {
              window.location.replace(GATEWAY_URL + '/?host=' + encodeURIComponent(host));
            } else {
              setState(
                '仍未能启动。请从开始菜单运行「OfficeAI」后重试。' +
                  (LAUNCH_COMMAND ? '（启动命令：' + LAUNCH_COMMAND + '）' : '')
              );
            }
          });
        };
      }
      setState('');
    });
  };
})();
