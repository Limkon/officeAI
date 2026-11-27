/* global wps, et, wpp, Application */
/**
 * OfficeAI WPS 加载项（本地离线版）
 * 被安装到 %APPDATA%\kingsoft\wps\jsaddons\officeai_1.0.0\ 下，WPS 用 file:// 直接加载
 * 点击按钮时：探活 -> 未启动则 ShellExecute 拉起 OfficeAI -> 打开侧边栏
 * 安装时 __OFFICEAI_*__ 占位符会被替换为真实值
 */
(function () {
  var GATEWAY_URL = '__OFFICEAI_GATEWAY_URL__';
  var LAUNCH_COMMAND = '__OFFICEAI_LAUNCH_COMMAND__';
  var LAUNCH_URL = '__OFFICEAI_LAUNCH_URL__';
  var BOOTSTRAP_URL = '__OFFICEAI_BOOTSTRAP_URL__';

  if (GATEWAY_URL.indexOf('__OFFICEAI_') === 0) {
    GATEWAY_URL = 'http://127.0.0.1:18080';
    LAUNCH_COMMAND = '';
    LAUNCH_URL = 'officeai://activate/' + encodeURIComponent(GATEWAY_URL + '/start?app=wps');
    BOOTSTRAP_URL = GATEWAY_URL + '/wps/bootstrap.html';
  }

  var TASKPANE_KEY = 'officeai_taskpane_id';
  var HEALTH_TIMEOUT_MS = 1200;
  var LAUNCH_POLL_MS = 1000;
  var LAUNCH_POLL_ATTEMPTS = 36;
  var launching = false;

  function getApp() {
    if (typeof window.Application !== 'undefined' && window.Application) return window.Application;
    if (typeof wps !== 'undefined' && wps.WpsApplication) return wps.WpsApplication();
    if (typeof et !== 'undefined' && et.EtApplication) return et.EtApplication();
    if (typeof wpp !== 'undefined' && wpp.WppApplication) return wpp.WppApplication();
    return null;
  }

  function detectWpsHostType() {
    var app = getApp();
    if (!app) return 'wps-writer';
    try {
      var name = String(app.Name || '').toLowerCase();
      if (name.indexOf('et') !== -1 || name.indexOf('spread') !== -1 || app.ActiveWorkbook) return 'wps-spreadsheets';
      if (name.indexOf('wpp') !== -1 || name.indexOf('present') !== -1 || app.ActivePresentation) return 'wps-presentation';
    } catch (e) {}
    return 'wps-writer';
  }

  function log(msg) {
    try { console.log('[OfficeAI-WPS] ' + msg); } catch (e) {}
  }

  // 立即给用户反馈：点按钮后 0ms 内弹轻提示，避免“没反应”错觉
  function toastStarting() {
    var app = getApp();
    // 用 ShowDialog 做一个极简的“正在启动”浮层，400ms 后会被 launcher 覆盖
    try {
      if (app && typeof app.ShowDialog === 'function') {
        var url = BOOTSTRAP_URL + '?app=' + encodeURIComponent(detectWpsHostType()) + '&toast=1';
        // 尝试用最小尺寸的 dialog 当 toast，失败则降级为 Alert
        // 不阻塞后续探活，独立 try/catch
        try { app.ShowDialog(url, 'OfficeAI 正在启动', 420, 200, false); } catch (e2) {}
      }
    } catch (e) {}
    log('button clicked, host=' + detectWpsHostType() + ' gateway=' + GATEWAY_URL + ' cmd=' + (LAUNCH_COMMAND || LAUNCH_URL));
  }

  // 多地址探活：依次尝试 127.0.0.1、localhost、https，任一成功即认为存活
  function gatewayAlive() {
    var candidates = [];
    if (GATEWAY_URL) candidates.push(GATEWAY_URL);
    if (GATEWAY_URL.indexOf('127.0.0.1:18080') !== -1) {
      candidates.push(GATEWAY_URL.replace('127.0.0.1:18080', 'localhost:18080'));
      candidates.push(GATEWAY_URL.replace('http://127.0.0.1:18080', 'https://127.0.0.1:18443'));
    } else if (GATEWAY_URL.indexOf('localhost:18080') !== -1) {
      candidates.push(GATEWAY_URL.replace('localhost:18080', '127.0.0.1:18080'));
    }
    // 去重
    var uniq = [];
    for (var i = 0; i < candidates.length; i++) if (uniq.indexOf(candidates[i]) === -1) uniq.push(candidates[i]);
    candidates = uniq;

    return new Promise(function (resolve) {
      var idx = 0;
      var timer = null;
      var settled = false;
      function finish(ok) { if (settled) return; settled = true; if (timer) clearTimeout(timer); resolve(ok); }
      timer = setTimeout(function () { finish(false); }, HEALTH_TIMEOUT_MS * Math.max(1, candidates.length));
      function tryOne() {
        if (idx >= candidates.length) { finish(false); return; }
        var base = candidates[idx++];
        var xhr = null;
        var done = false;
        var to = setTimeout(function () { if (!done) { try { xhr && xhr.abort(); } catch (e) {} tryOne(); } }, HEALTH_TIMEOUT_MS);
        try {
          xhr = new XMLHttpRequest();
          xhr.open('GET', base + '/api/health?_t=' + Date.now(), true);
          xhr.onreadystatechange = function () {
            if (xhr.readyState !== 4 || done) return;
            done = true;
            clearTimeout(to);
            if (xhr.status >= 200 && xhr.status < 300) { finish(true); } else { tryOne(); }
          };
          xhr.onerror = function () { if (done) return; done = true; clearTimeout(to); tryOne(); };
          xhr.send();
        } catch (e) { clearTimeout(to); tryOne(); }
      }
      tryOne();
    });
  }

  function showLaunchFailed(hostType) {
    var app = getApp();
    var cmd = LAUNCH_COMMAND || LAUNCH_URL || 'OfficeAI.exe';
    var msg =
      'OfficeAI 本地服务启动失败。\n\n' +
      '可能原因：\n' +
      '1. WPS 安全设置拦截了外部程序（首次会弹确认，请点“允许”）\n' +
      '2. WPS → 设置 → 安全 → 宏/加载项被禁用\n' +
      '3. 安装路径不存在或端口 18080 被占用\n\n' +
      '请尝试：\n' +
      '• 按 Win 键搜索“OfficeAI”手动运行，然后重试按钮\n' +
      '• 重启 WPS 后再点一次\n' +
      '• 检查開始選單是否有“OfficeAI”快捷方式\n\n' +
      '启动命令：' + cmd + '\n网关地址：' + GATEWAY_URL;
    try { if (app && typeof app.Alert === 'function') { app.Alert(msg); return; } } catch (e) {}
    try { if (typeof wps !== 'undefined' && typeof wps.Alert === 'function') { wps.Alert(msg); return; } } catch (e) {}
    try { alert(msg); } catch (e2) {}
    try {
      if (app && typeof app.ShowDialog === 'function') {
        var fallbackUrl = BOOTSTRAP_URL + '?app=' + encodeURIComponent(hostType || 'wps-writer');
        app.ShowDialog(fallbackUrl, 'OfficeAI 启动失败 - 手动启动指引', 760, 580, false);
      }
    } catch (e3) {}
  }

  function shellExecuteWithFallback(cmd) {
    var errs = [];
    // 1) 直接 exe 路径
    try {
      if (typeof wps !== 'undefined' && wps.OAAssist && typeof wps.OAAssist.ShellExecute === 'function') {
        wps.OAAssist.ShellExecute(cmd);
        log('ShellExecute via wps.OAAssist: ' + cmd);
        return true;
      }
    } catch (e) { errs.push('wps.OAAssist:' + (e.message || e)); }
    var app = getApp();
    try {
      if (app && app.OAAssist && typeof app.OAAssist.ShellExecute === 'function') {
        app.OAAssist.ShellExecute(cmd);
        log('ShellExecute via app.OAAssist: ' + cmd);
        return true;
      }
    } catch (e) { errs.push('app.OAAssist:' + (e.message || e)); }
    // 2) 尝试 cmd /c start 包裹（处理路径含空格）
    try {
      var quoted = cmd.indexOf(' ') !== -1 && cmd[0] !== '"' ? '"' + cmd + '"' : cmd;
      var wrapper = 'cmd.exe /c start "" ' + quoted;
      if (typeof wps !== 'undefined' && wps.OAAssist && typeof wps.OAAssist.ShellExecute === 'function') {
        wps.OAAssist.ShellExecute(wrapper);
        log('ShellExecute wrapper: ' + wrapper);
        return true;
      }
    } catch (e) { errs.push('wrapper:' + (e.message || e)); }
    log('ShellExecute all failed for ' + cmd + ' errs=' + errs.join(' | '));
    return false;
  }

  function launchGateway(actionId, hostType) {
    var app = getApp();
    var targetUrl = BOOTSTRAP_URL + '?app=' + encodeURIComponent(hostType) + (actionId ? '&action=' + encodeURIComponent(actionId) : '');
    var cmd = LAUNCH_COMMAND;
    var url = LAUNCH_URL;
    // 立即给用户可视反馈：弹出启动页
    try {
      if (app && typeof app.ShowDialog === 'function' && BOOTSTRAP_URL.indexOf('http') === 0) {
        app.ShowDialog(targetUrl, 'OfficeAI 正在启动…', 760, 560, false);
      }
    } catch (e) {}

    var started = false;
    if (cmd) {
      started = shellExecuteWithFallback(cmd);
      if (!started && url) {
        log('exe launch failed, trying protocol URL: ' + url);
        started = shellExecuteWithFallback(url);
      }
    } else if (url) {
      started = shellExecuteWithFallback(url);
    } else {
      showLaunchFailed(hostType);
      return false;
    }

    if (!started) {
      log('launchGateway failed, both exe and protocol failed');
      // 不立即弹失败，让 waitForGateway 去弹，避免重复
    } else {
      log('launchGateway ShellExecute dispatched, waiting for health...');
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

  function openOrUpdateTaskPane(actionId) {
    var app = getApp();
    if (!app) return;
    var hostType = detectWpsHostType();
    var targetUrl = GATEWAY_URL + '/?host=wps&app=' + encodeURIComponent(hostType) + (actionId ? '&action=' + encodeURIComponent(actionId) : '') + '&t=' + Date.now();
    log('openOrUpdateTaskPane ' + targetUrl);
    try {
      var tsId = null;
      if (app.PluginStorage && typeof app.PluginStorage.getItem === 'function') tsId = app.PluginStorage.getItem(TASKPANE_KEY);
      if (tsId && typeof app.GetTaskPane === 'function') {
        var existingPane = app.GetTaskPane(tsId);
        if (existingPane) {
          existingPane.Visible = true;
          if (actionId && existingPane.Navigate) existingPane.Navigate(targetUrl);
          return;
        }
      }
    } catch (e) { log('reuse pane failed: ' + e); }
    try {
      if (typeof app.CreateTaskPane === 'function') {
        var pane = app.CreateTaskPane(targetUrl, 'OfficeAI 智能助手');
        if (pane) {
          pane.DockPosition = 2;
          pane.Width = 400;
          pane.Visible = true;
          if (app.PluginStorage && typeof app.PluginStorage.setItem === 'function') app.PluginStorage.setItem(TASKPANE_KEY, pane.ID);
          return;
        }
      }
    } catch (err) { log('CreateTaskPane failed: ' + err); }
    try {
      if (typeof app.ShowDialog === 'function') app.ShowDialog(targetUrl, 'OfficeAI 智能助手', 960, 680, false);
    } catch (e) { log('ShowDialog fallback failed: ' + e); }
  }

  function openDashboard() {
    var app = getApp();
    var dashUrl = GATEWAY_URL + '/dashboard';
    log('openDashboard ' + dashUrl);
    if (app && typeof app.ShowDialog === 'function') app.ShowDialog(dashUrl, 'OfficeAI 模型与集成配置中心', 1100, 760, false);
    else openOrUpdateTaskPane('');
  }

  function ensureGatewayThen(actionId) {
    var hostType = detectWpsHostType();
    toastStarting();
    return gatewayAlive().then(function (ok) {
      if (ok) {
        log('gateway already alive, open pane');
        try { openOrUpdateTaskPane(actionId); } catch (e) { log('open pane failed: ' + e); }
        return true;
      }
      if (launching) {
        try { var a = getApp(); if (a && a.Alert) a.Alert('OfficeAI 正在启动中，请稍等 2-3 秒后重试…'); } catch (e) {}
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
          log('gateway started, open pane');
          try { openOrUpdateTaskPane(actionId); } catch (e) {}
          return true;
        }
        log('waitForGateway timeout');
        showLaunchFailed(hostType);
        return false;
      });
    }).catch(function (err) {
      launching = false;
      log('ensureGatewayThen catch: ' + err);
      showLaunchFailed(detectWpsHostType());
      return false;
    });
  }

  window.OnAddinLoad = function (ribbonUI) {
    window.OfficeAIRibbonUI = ribbonUI;
    log('OnAddinLoad, gateway=' + GATEWAY_URL);
    gatewayAlive().then(function (ok) { log('probe on load: ' + ok); });
    return true;
  };
  window.OnGetEnabled = function () { return true; };
  window.GetImage = function () {
    // 返回 http 图标，若网关未启动 WPS 会显示占位，但不影响按钮可用
    return GATEWAY_URL + '/wps/assets/icon-32.png';
  };
  window.OnAction = function (control) {
    try {
      log('OnAction ' + (control && control.Id));
      var id = control && control.Id ? control.Id : 'btnOpenTaskPane';
      if (id === 'btnOpenDashboard') {
        ensureGatewayThen('').then(function (ok) {
          if (ok) try { openDashboard(); } catch (e) { showLaunchFailed(detectWpsHostType()); }
        }).catch(function () { showLaunchFailed(detectWpsHostType()); });
        return;
      }
      var actionMap = { btnOpenTaskPane: '', btnPolish: 'polish', btnTranslate: 'translate-en', btnSummarize: 'summarize', btnFormula: 'excel-formula', btnPptOutline: 'ppt-outline' };
      ensureGatewayThen(actionMap[id] || '').catch(function () { showLaunchFailed(detectWpsHostType()); });
    } catch (err) {
      log('OnAction exception: ' + err);
      showLaunchFailed(detectWpsHostType());
    }
  };

  window.OfficeAIBootstrap = function (hostParam) {
    var host = hostParam === 'wps' ? detectWpsHostType() : hostParam || 'standalone';
    gatewayAlive().then(function (ok) {
      if (ok) { window.location.replace(GATEWAY_URL + '/?host=' + encodeURIComponent(host)); return; }
      var shell = document.body || document.documentElement;
      shell.setAttribute('style', 'margin:0;font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;background:#f8fafc;color:#0f172a;display:flex;align-items:center;justify-content:center;min-height:100vh;');
      shell.innerHTML =
        '<div style="max-width:460px;padding:24px">' +
        '<div style="background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:20px">' +
        '<h1 style="font-size:17px;margin:0 0 8px">OfficeAI 服务未启动</h1>' +
        '<p style="font-size:13px;line-height:1.7;color:#475569;margin:0 0 14px">点击下方按钮会在需要时自动拉起 OfficeAI 本地服务，随后自动进入助手。</p>' +
        '<button id="officeai-start" style="width:100%;border:0;border-radius:10px;padding:11px 14px;font-size:14px;font-weight:600;color:#fff;background:#2563eb;cursor:pointer">一键启动 OfficeAI</button>' +
        '<div id="officeai-state" style="font-size:12px;color:#64748b;margin-top:12px"></div>' +
        '<p style="font-size:11px;color:#94a3b8;margin-top:12px">启动命令：' + (LAUNCH_COMMAND || LAUNCH_URL || 'OfficeAI.exe') + '<br>网关：' + GATEWAY_URL + '</p>' +
        '</div></div>';
      var stateEl = document.getElementById('officeai-state');
      function setState(t) { if (stateEl) stateEl.textContent = t; }
      var btn = document.getElementById('officeai-start');
      if (btn) btn.onclick = function () {
        setState('正在唤起 OfficeAI…');
        launchGateway('', host);
        waitForGateway().then(function (started) {
          if (started) window.location.replace(GATEWAY_URL + '/?host=' + encodeURIComponent(host));
          else setState('仍未能启动。请从开始菜单运行「OfficeAI」后重试。' + (LAUNCH_COMMAND ? '（' + LAUNCH_COMMAND + '）' : ''));
        });
      };
      setState('');
    });
  };
})();
