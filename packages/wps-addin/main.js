/* global wps, et, wpp, Application */
// 最小兜底：保证即使后续脚本异常，按钮也能弹框证明链路已通
var __OA_G = (typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : typeof self !== "undefined" ? self : this);
try {
  var __oaFallback = function (c) { try { var id=(c&&c.Id)||"unknown"; if(typeof wps!=="undefined"&&wps.Alert) wps.Alert("OfficeAI 调试: 按钮已触发 "+id+"\n(若只看到此框，说明后续脚本加载异常，请点‘调试日志’或重装)"); else if(typeof alert!=="undefined") alert("OfficeAI 调试: 按钮已触发 "+id); } catch(e){} };
  __OA_G.OnAction = __OA_G.OnAction || __oaFallback;
  __OA_G.OnAddinLoad = __OA_G.OnAddinLoad || function(){ return true; };
  __OA_G.OnAddInLoad = __OA_G.OnAddInLoad || __OA_G.OnAddinLoad;
  __OA_G.OnGetEnabled = __OA_G.OnGetEnabled || function(){ return true; };
  __OA_G.GetImage = __OA_G.GetImage || function(){ return ""; };
  try { if (typeof OnAction === "undefined") OnAction = __OA_G.OnAction; } catch(e){}
  try { if (typeof OnAddinLoad === "undefined") OnAddinLoad = __OA_G.OnAddinLoad; } catch(e){}
  try { if (typeof OnAddInLoad === "undefined") OnAddInLoad = __OA_G.OnAddInLoad; } catch(e){}
  try { if (typeof OnGetEnabled === "undefined") OnGetEnabled = __OA_G.OnGetEnabled; } catch(e){}
  try { if (typeof GetImage === "undefined") GetImage = __OA_G.GetImage; } catch(e){}
} catch(e) {}
/**
 * OfficeAI WPS 加载项（本地离线版）- 调试增强版
 * 安装到 %APPDATA%\kingsoft\wps\jsaddons\officeai_1.0.0\ 下
 * 点击按钮：探活 -> 未启动则 ShellExecute 拉起 -> 打开侧边栏
 * 全量步骤均写入 DEBUG_LOG，点“调试日志”可查看
 */
(function () {
  try {
  var __OA_G2 = (typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : typeof self !== "undefined" ? self : this);
  if (typeof __OA_G === "undefined") { try { __OA_G = __OA_G2; } catch(e){} }
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
  // 纯静态 UI：完全脱离 127.0.0.1，任务窗格走 file:// 本地离线包（参考 it235 的 VSTO 无网关直连）
  var USE_STATIC_UI = true;
  function getStaticUiUrl(query) {
    try {
      // 优先用 wps.GetUrlPath()（WPS 标准），否则用 location.href 推导
      var baseDir = '';
      try {
        if (typeof wps !== 'undefined' && wps && typeof wps.GetUrlPath === 'function') {
          var gp = wps.GetUrlPath();
          if (gp) {
            baseDir = gp;
            if (baseDir.slice(-1) !== '/' && baseDir.slice(-1) !== '\\') baseDir += '/';
            dbg('GetUrlPath', baseDir);
          }
        }
      } catch(e) { dbg('GetUrlPath 失败', e.message||e); }
      if (!baseDir) {
        var loc = (typeof location !== 'undefined' && location.href) ? location.href : '';
        if (loc) {
          baseDir = loc.substring(0, loc.lastIndexOf('/')+1);
          dbg('location 推导 baseDir', baseDir);
        }
      }
      if (!baseDir) baseDir = 'file:///C:/Users/Public/OfficeAI/';
      // 统一为 file:/// 前缀
      if (baseDir.indexOf('file://') !== 0 && baseDir.indexOf('http') !== 0) {
        if (baseDir.match(/^[A-Za-z]:/)) baseDir = 'file:///' + baseDir.replace(/\\/g,'/');
        else if (baseDir[0] === '/') baseDir = 'file://' + baseDir;
      }
      var url = baseDir + 'ui/index.html' + query;
      // 修正 file:/// 前缀（避免 file://C:）并清理多余斜杠
      if (url.indexOf('file://') === 0) {
        url = url.replace('file:////', 'file:///');
        if (url.indexOf('file://C:') === 0 || url.indexOf('file://c:') === 0) {
          url = 'file:///' + url.substring(7);
        }
        // file:/// 后的双斜杠仅保留一处
        var prefix = 'file:///';
        var rest = url.substring(prefix.length).replace(/\/\/+/g, '/');
        url = prefix + rest;
      }
      dbg('静态 UI URL', url);
      return url;
    } catch(e) { dbg('getStaticUiUrl 失败', e.message||e); return GATEWAY_URL + '/?host=wps' + query; }
  }

  // ===================== 调试基础设施 =====================
  var DEBUG_LOG = [];
  var DEBUG_START = Date.now();
  function dbg(msg, data) {
    var elapsed = ((Date.now() - DEBUG_START) / 1000).toFixed(3) + 's';
    var line = '[' + elapsed + '] ' + msg;
    if (data !== undefined) {
      try { line += ' | ' + (typeof data === 'string' ? data : JSON.stringify(data)); } catch (e) { line += ' | ' + String(data); }
    }
    DEBUG_LOG.push(line);
    if (DEBUG_LOG.length > 300) DEBUG_LOG.shift();
    try { console.log('[OfficeAI-WPS-DBG] ' + line); } catch (e) {}
  }
  function getDebugText() {
    return 'OfficeAI WPS 调试日志\n时间: ' + new Date().toISOString() + '\n网关: ' + GATEWAY_URL + '\n启动命令: ' + (LAUNCH_COMMAND || '(空)') + '\n启动URL: ' + (LAUNCH_URL || '(空)') + '\n宿主: ' + detectWpsHostType() + '\n--- 步骤 ---\n' + DEBUG_LOG.join('\n');
  }
  function showDebugLog() {
    var text = getDebugText();
    dbg('显示调试日志');
    var app = getApp();
    // 优先用 ShowDialog 展示 HTML
    try {
      if (app && typeof app.ShowDialog === 'function') {
        var html = '<html><head><meta charset="utf-8"><style>body{font-family:Consolas,monospace;font-size:11px;padding:16px;white-space:pre-wrap;word-break:break-all}h2{font-size:13px}button{margin-top:12px;padding:8px 16px}</style></head><body><h2>OfficeAI WPS 调试日志</h2><pre>' + text.replace(/</g,'&lt;').replace(/>/g,'&gt;') + '</pre><button onclick="navigator.clipboard&&navigator.clipboard.writeText(document.querySelector(\'pre\').innerText);alert(\'已复制\')">复制日志</button></body></html>';
        var dataUrl = 'data:text/html;charset=utf-8,' + encodeURIComponent(html);
        app.ShowDialog(dataUrl, 'OfficeAI 调试日志', 900, 640, false);
        return;
      }
    } catch (e) { dbg('ShowDialog 调试失败', e.message || e); }
    try {
      if (app && typeof app.Alert === 'function') { app.Alert(text.slice(0, 4000)); return; }
    } catch (e) {}
    try { alert(text.slice(0, 4000)); } catch (e) {}
  }
  // 全局暴露，便于控制台手动调用
  try { window.OfficeAIDebugLog = DEBUG_LOG; window.OfficeAIShowDebug = showDebugLog; window.OfficeAIDbg = dbg; } catch (e) {}

  dbg('初始化完成', { GATEWAY_URL: GATEWAY_URL, LAUNCH_COMMAND: LAUNCH_COMMAND, LAUNCH_URL: LAUNCH_URL, BOOTSTRAP_URL: BOOTSTRAP_URL, location: (typeof location !== 'undefined' ? location.href : 'unknown') });

  // ===================== 宿主探测 =====================
  function getApp() {
    // JS 插件的文档对象：优先返回 Window/Application，再回退到 wps.WpsApplication 的 COM 对象
    // 注意：CreateTaskPane / ShowDialog / PluginStorage 并不在 WpsApplication 上，而在 wps / Application 全局上
    try { if (typeof Application !== 'undefined' && Application) return Application; } catch(e){}
    var info = [];
    if (typeof window !== 'undefined' && typeof window.Application !== 'undefined' && window.Application) { info.push('window.Application'); return window.Application; }
    try { if (typeof wps !== 'undefined' && wps.WpsApplication) { var a = wps.WpsApplication(); info.push('wps.WpsApplication'); try{ dbg('getApp via wps', String(a && a.Name)); }catch(e){} return a; } } catch (e) { dbg('getApp wps failed', e.message||e); }
    try { if (typeof et !== 'undefined' && et.EtApplication) { var b = et.EtApplication(); info.push('et'); return b; } } catch (e) {}
    try { if (typeof wpp !== 'undefined' && wpp.WppApplication) { var c = wpp.WppApplication(); info.push('wpp'); return c; } } catch (e) {}
    dbg('getApp 未找到宿主', info.join(',')||'none');
    return null;
  }
  // 兼容 WPS 标准示例：wps.CreateTaskPane / wps.ShowDialog / Application.CreateTaskPane
  function getTaskPaneHosts() {
    var list = [];
    var seen = {};
    function push(o, name) {
      if (!o) return;
      var key = name + ':' + (o === __OA_G ? 'g' : String(o));
      if (seen[key]) return;
      seen[key]=1;
      list.push({ obj: o, name: name });
    }
    try { if (typeof wps !== 'undefined' && wps) push(wps, 'wps'); } catch(e){}
    try { if (typeof Application !== 'undefined' && Application) push(Application, 'Application'); } catch(e){}
    try { if (typeof window !== 'undefined' && window.wps) push(window.wps, 'window.wps'); } catch(e){}
    try { if (typeof window !== 'undefined' && window.Application) push(window.Application, 'window.Application'); } catch(e){}
    try { if (typeof self !== 'undefined' && self.wps) push(self.wps, 'self.wps'); } catch(e){}
    try { var a = getApp(); if (a) push(a, 'getApp()'); } catch(e){}
    // 某些 WPS 版本把 API 挂在 __OA_G
    try { if (__OA_G && __OA_G.wps) push(__OA_G.wps, '__OA_G.wps'); } catch(e){}
    return list;
  }
  function dbgHosts(tag, hosts) {
    try {
      var s = hosts.map(function(h){ var hasCT = !!(h.obj && typeof h.obj.CreateTaskPane === 'function'); var hasSD = !!(h.obj && typeof h.obj.ShowDialog === 'function'); var hasPS = !!(h.obj && h.obj.PluginStorage); return h.name + (hasCT?'[CT]':'') + (hasSD?'[SD]':'') + (hasPS?'[PS]':''); }).join(', ');
      dbg(tag, s || 'none');
    } catch(e){}
  }

  function detectWpsHostType() {
    var app = getApp();
    if (!app) { dbg('detectHost: 无 app，默认 writer'); return 'wps-writer'; }
    try {
      var name = String(app.Name || '').toLowerCase();
      dbg('detectHost app.Name=' + name + ' has ActiveWorkbook=' + !!app.ActiveWorkbook + ' has ActivePresentation=' + !!app.ActivePresentation);
      if (name.indexOf('et') !== -1 || name.indexOf('spread') !== -1 || app.ActiveWorkbook) return 'wps-spreadsheets';
      if (name.indexOf('wpp') !== -1 || name.indexOf('present') !== -1 || app.ActivePresentation) return 'wps-presentation';
    } catch (e) { dbg('detectHost 异常', e.message||e); }
    return 'wps-writer';
  }

  // ===================== 健康检查（多候选） =====================
  function gatewayAlive() {
    var candidates = [];
    if (GATEWAY_URL) candidates.push(GATEWAY_URL);
    if (GATEWAY_URL.indexOf('127.0.0.1:18080') !== -1) {
      candidates.push(GATEWAY_URL.replace('127.0.0.1:18080', 'localhost:18080'));
      candidates.push(GATEWAY_URL.replace('http://127.0.0.1:18080', 'https://127.0.0.1:18443'));
    } else if (GATEWAY_URL.indexOf('localhost:18080') !== -1) {
      candidates.push(GATEWAY_URL.replace('localhost:18080', '127.0.0.1:18080'));
    }
    var uniq = [];
    for (var i = 0; i < candidates.length; i++) if (uniq.indexOf(candidates[i]) === -1) uniq.push(candidates[i]);
    candidates = uniq;
    dbg('gatewayAlive 开始', candidates.join(' | '));

    return new Promise(function (resolve) {
      var idx = 0;
      var timer = null;
      var settled = false;
      function finish(ok, reason) {
        if (settled) return;
        settled = true;
        if (timer) clearTimeout(timer);
        dbg('gatewayAlive 结束', (ok ? '成功' : '失败') + (reason ? ' (' + reason + ')' : ''));
        resolve(ok);
      }
      timer = setTimeout(function () { finish(false, '总超时 ' + (HEALTH_TIMEOUT_MS * candidates.length) + 'ms'); }, HEALTH_TIMEOUT_MS * Math.max(1, candidates.length) + 300);
      function tryOne() {
        if (idx >= candidates.length) { finish(false, '全部候选失败'); return; }
        var base = candidates[idx++];
        var url = base + '/api/health?_t=' + Date.now();
        var xhr = null;
        var done = false;
        var to = setTimeout(function () { if (!done) { dbg('XHR 超时', url); try { xhr && xhr.abort(); } catch (e) {} tryOne(); } }, HEALTH_TIMEOUT_MS);
        dbg('XHR 尝试', url);
        try {
          xhr = new XMLHttpRequest();
          xhr.open('GET', url, true);
          xhr.onreadystatechange = function () {
            if (xhr.readyState !== 4 || done) return;
            done = true;
            clearTimeout(to);
            var status = xhr.status;
            var body = ''; try { body = (xhr.responseText || '').slice(0, 200); } catch (e) {}
            dbg('XHR 响应', base + ' status=' + status + ' body=' + body);
            if (status >= 200 && status < 300) { finish(true, base); } else { tryOne(); }
          };
          xhr.onerror = function () { if (done) return; done = true; clearTimeout(to); dbg('XHR onerror', url); tryOne(); };
          xhr.send();
        } catch (e) { clearTimeout(to); dbg('XHR 异常', e.message || e); tryOne(); }
      }
      tryOne();
    });
  }

  function showLaunchFailed(hostType) {
    var app = getApp();
    var cmd = LAUNCH_COMMAND || LAUNCH_URL || 'OfficeAI.exe';
    var logTail = DEBUG_LOG.slice(-80).join('\n');
    var msg =
      'OfficeAI 本地服务启动失败。\n\n' +
      '启动命令：' + cmd + '\n网关地址：' + GATEWAY_URL + '\n宿主：' + hostType + '\n\n' +
      '请尝试：\n' +
      '• 按 Win 键搜索“OfficeAI”手动运行，然后重试按钮\n' +
      '• 重启 WPS 后再点一次\n' +
      '• 点功能区“调试日志”查看详细步骤并反馈\n\n' +
      '--- 最近日志 ---\n' + logTail;
    dbg('showLaunchFailed', msg.slice(0, 500));
    try { if (app && typeof app.Alert === 'function') { app.Alert(msg.slice(0, 3800)); return; } } catch (e) { dbg('Alert 失败', e.message||e); }
    try { if (typeof wps !== 'undefined' && typeof wps.Alert === 'function') { wps.Alert(msg.slice(0, 3800)); return; } } catch (e) {}
    try { alert(msg.slice(0, 3800)); } catch (e2) {}
    try {
      var fallbackUrl = BOOTSTRAP_URL + '?app=' + encodeURIComponent(hostType || 'wps-writer');
      var hs2 = getTaskPaneHosts();
      for (var jj=0; jj<hs2.length; jj++) {
        var hh3 = hs2[jj].obj;
        if (hh3 && typeof hh3.ShowDialog === 'function') {
          dbg('ShowDialog fallback @'+hs2[jj].name, fallbackUrl);
          try { hh3.ShowDialog(fallbackUrl, 'OfficeAI 启动失败 - 调试', 760, 580, false); break; } catch(e){ dbg('ShowDialog fallback 失败 @'+hs2[jj].name, e.message||e); }
        }
      }
    } catch (e3) { dbg('ShowDialog fallback 失败', e3.message||e3); }
  }

  function shellExecuteWithFallback(cmd) {
    var errs = [];
    dbg('ShellExecute 尝试', cmd);
    try {
      if (typeof wps !== 'undefined' && wps.OAAssist && typeof wps.OAAssist.ShellExecute === 'function') {
        dbg('尝试 wps.OAAssist.ShellExecute');
        wps.OAAssist.ShellExecute(cmd);
        dbg('wps.OAAssist 成功');
        return true;
      } else { errs.push('wps.OAAssist 不存在'); dbg('wps.OAAssist 不存在'); }
    } catch (e) { var m = e.message||e; errs.push('wps.OAAssist:'+m); dbg('wps.OAAssist 失败', m); }
    var app = getApp();
    try {
      if (app && app.OAAssist && typeof app.OAAssist.ShellExecute === 'function') {
        dbg('尝试 app.OAAssist.ShellExecute');
        app.OAAssist.ShellExecute(cmd);
        dbg('app.OAAssist 成功');
        return true;
      } else { errs.push('app.OAAssist 不存在'); dbg('app.OAAssist 不存在'); }
    } catch (e) { var m2 = e.message||e; errs.push('app.OAAssist:'+m2); dbg('app.OAAssist 失败', m2); }
    try {
      var quoted = cmd.indexOf(' ') !== -1 && cmd[0] !== '"' ? '"' + cmd + '"' : cmd;
      var wrapper = 'cmd.exe /c start "" ' + quoted;
      dbg('尝试 wrapper', wrapper);
      if (typeof wps !== 'undefined' && wps.OAAssist && typeof wps.OAAssist.ShellExecute === 'function') {
        wps.OAAssist.ShellExecute(wrapper);
        dbg('wrapper 成功');
        return true;
      }
    } catch (e) { var m3 = e.message||e; errs.push('wrapper:'+m3); dbg('wrapper 失败', m3); }
    dbg('ShellExecute 全部失败', errs.join(' | '));
    return false;
  }

  function launchGateway(actionId, hostType) {
    var app = getApp();
    var targetUrl = BOOTSTRAP_URL + '?app=' + encodeURIComponent(hostType) + (actionId ? '&action=' + encodeURIComponent(actionId) : '');
    var cmd = LAUNCH_COMMAND;
    var url = LAUNCH_URL;
    dbg('launchGateway', { cmd: cmd, url: url, targetUrl: targetUrl, hostType: hostType });
    try {
      if (BOOTSTRAP_URL.indexOf('http') === 0) {
        var hs = getTaskPaneHosts();
        for (var ii=0; ii<hs.length; ii++) {
          var hh2 = hs[ii].obj;
          if (hh2 && typeof hh2.ShowDialog === 'function') {
            dbg('ShowDialog 启动页 @'+hs[ii].name, targetUrl);
            try { hh2.ShowDialog(targetUrl, 'OfficeAI 正在启动…', 760, 560, false); break; } catch(e){ dbg('ShowDialog 启动页失败 @'+hs[ii].name, e.message||e); }
          }
        }
      }
    } catch (e) { dbg('ShowDialog 启动页失败', e.message||e); }

    var started = false;
    if (cmd) {
      started = shellExecuteWithFallback(cmd);
      if (!started && url) {
        dbg('exe 启动失败，尝试协议 URL');
        started = shellExecuteWithFallback(url);
      }
    } else if (url) {
      started = shellExecuteWithFallback(url);
    } else {
      dbg('无启动命令，放弃');
      showLaunchFailed(hostType);
      return false;
    }
    dbg('launchGateway 结果', started ? '已派发' : '全部失败');
    return started;
  }

  function waitForGateway() {
    var attempt = 0;
    dbg('waitForGateway 开始', '最多 ' + LAUNCH_POLL_ATTEMPTS + ' 次');
    return new Promise(function (resolve) {
      function step() {
        gatewayAlive().then(function (ok) {
          dbg('wait 轮询 #' + (attempt+1) + ' ' + (ok ? '成功' : '未就绪'));
          if (ok) return resolve(true);
          attempt += 1;
          if (attempt >= LAUNCH_POLL_ATTEMPTS) { dbg('waitForGateway 超时'); return resolve(false); }
          setTimeout(step, LAUNCH_POLL_MS);
        });
      }
      step();
    });
  }

  function openOrUpdateTaskPane(actionId) {
    var hostType = detectWpsHostType();
    var targetUrl = USE_STATIC_UI
      ? getStaticUiUrl('?host=wps&app=' + encodeURIComponent(hostType) + (actionId ? '&action=' + encodeURIComponent(actionId) : '') + '&t=' + Date.now())
      : GATEWAY_URL + '/?host=wps&app=' + encodeURIComponent(hostType) + (actionId ? '&action=' + encodeURIComponent(actionId) : '') + '&t=' + Date.now();
    dbg('openOrUpdateTaskPane' + (USE_STATIC_UI ? ' [静态 file://]' : ' [网关]'), targetUrl);
    var hosts = getTaskPaneHosts();
    dbgHosts('可用宿主', hosts);
    // 1) 尝试复用已有 pane（遍历所有宿主的 PluginStorage / GetTaskPane）
    for (var hi=0; hi<hosts.length; hi++) {
      var h = hosts[hi].obj;
      var hname = hosts[hi].name;
      try {
        if (h && h.PluginStorage && typeof h.PluginStorage.getItem === 'function') {
          var tsId = h.PluginStorage.getItem(TASKPANE_KEY);
          if (tsId) {
            dbg('PluginStorage get @'+hname, String(tsId));
            if (typeof h.GetTaskPane === 'function') {
              var existingPane = h.GetTaskPane(tsId);
              dbg('GetTaskPane @'+hname, String(!!existingPane));
              if (existingPane) {
                try { existingPane.Visible = true; } catch(e){}
                dbg('复用已有 pane 并显示 @'+hname);
                if (actionId && existingPane.Navigate) { try{ dbg('Navigate @'+hname, targetUrl); existingPane.Navigate(targetUrl); }catch(e){ dbg('Navigate 失败 @'+hname, e.message||e); } }
                return;
              }
            }
          }
        }
      } catch (e) { dbg('复用 pane 失败 @'+hname, e.message||e); }
    }
    // 2) 尝试新建 CreateTaskPane（按优先级遍历）
    for (var ci=0; ci<hosts.length; ci++) {
      var hh = hosts[ci].obj;
      var nn = hosts[ci].name;
      try {
        if (hh && typeof hh.CreateTaskPane === 'function') {
          dbg('CreateTaskPane 尝试 @'+nn, targetUrl);
          var pane = hh.CreateTaskPane(targetUrl);
          // 部分 WPS 版本需要第二个参数标题，部分不需要；兼容两种
          if (!pane && typeof hh.CreateTaskPane === 'function') {
            try { pane = hh.CreateTaskPane(targetUrl, 'OfficeAI 智能助手'); } catch(e){}
          }
          dbg('CreateTaskPane 结果 @'+nn, String(!!pane) + (pane ? ' ID=' + (pane.ID||pane.Id||'') : ''));
          if (pane) {
            try { pane.DockPosition = 2; } catch(e){}
            try { pane.Width = 400; } catch(e){}
            try { pane.Visible = true; } catch(e){}
            dbg('pane 已显示 @'+nn);
            for (var pi=0; pi<hosts.length; pi++) {
              var ph = hosts[pi].obj;
              try { if (ph && ph.PluginStorage && typeof ph.PluginStorage.setItem === 'function') { ph.PluginStorage.setItem(TASKPANE_KEY, pane.ID||pane.Id); dbg('已保存 pane ID @'+hosts[pi].name, pane.ID||pane.Id); break; } } catch(e){}
            }
            return;
          }
        } else {
          dbg('CreateTaskPane 不存在 @'+nn);
        }
      } catch (err) { dbg('CreateTaskPane 异常 @'+nn, err.message||err); }
    }
    // 3) 降级 ShowDialog（遍历）
    for (var di=0; di<hosts.length; di++) {
      var dh = hosts[di].obj;
      var dn = hosts[di].name;
      try {
        if (dh && typeof dh.ShowDialog === 'function') {
          dbg('ShowDialog 降级 @'+dn, targetUrl);
          dh.ShowDialog(targetUrl, 'OfficeAI 智能助手', 960, 680, false);
          dbg('ShowDialog 已调用 @'+dn);
          return;
        }
      } catch (e) { dbg('ShowDialog 降级失败 @'+dn, e.message||e); }
    }
    dbg('所有方式均失败，弹窗提示');
    try {
      var allNames = hosts.map(function(x){return x.name}).join(', ');
      var msg = 'OfficeAI 侧边栏创建失败：\nCreateTaskPane / ShowDialog 在当前宿主上均不可用。\n宿主：'+hostType+'\n可用宿主：'+allNames+'\n\n请重启 WPS 后重试，或点"调试日志"反馈。\n网关：'+GATEWAY_URL+'\n地址：'+targetUrl;
      if (typeof wps !== 'undefined' && wps && typeof wps.Alert === 'function') wps.Alert(msg);
      else if (typeof alert !== 'undefined') alert(msg);
    } catch(e){}
  }

  function openDashboard() {
    var dashUrl;
    if (USE_STATIC_UI) {
      // 优先 http 网关（设置页在 file:// 下也可能空白），失败再 file://
      // 这里做同步判断：若最近一次 gatewayAlive 成功则用 http，否则 file://
      // 简化：先尝试用 http，若打不开再由用户重试 file://（getStaticUiUrl）
      // 为兼容当前同步调用，先用 gatewayAlive 的最近结果（若无则假设存活）
      // 实际以 openOrUpdateTaskPane 的混合策略为准：先 http
      dashUrl = GATEWAY_URL + '/dashboard';
      // 探测是否可用，失败则回退 file://（异步探测后若失败再开 file:// 弹窗）
      try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', GATEWAY_URL + '/api/health?_t=' + Date.now(), false);
        xhr.send();
        if (xhr.status < 200 || xhr.status >= 300) throw new Error('no gateway');
        dbg('openDashboard [网关 http]', dashUrl);
      } catch(e) {
        dashUrl = getStaticUiUrl('?host=wps&mode=dashboard&t=' + Date.now());
        dbg('openDashboard 回退 [静态 file://]', dashUrl);
      }
    } else {
      dashUrl = GATEWAY_URL + '/dashboard';
      dbg('openDashboard', dashUrl);
    }
    var hosts = getTaskPaneHosts();
    for (var i=0;i<hosts.length;i++) {
      var h = hosts[i].obj;
      if (h && typeof h.ShowDialog === 'function') {
        dbg('ShowDialog dashboard @'+hosts[i].name, dashUrl);
        try { h.ShowDialog(dashUrl, 'OfficeAI 模型与集成配置中心', 1100, 760, false); return; } catch(e){ dbg('ShowDialog dashboard 失败 @'+hosts[i].name, e.message||e); }
      }
    }
    dbg('无 ShowDialog，降级到 taskpane');
    openOrUpdateTaskPane('');
  }

  function ensureGatewayThen(actionId) {
    var hostType = detectWpsHostType();
    dbg('ensureGatewayThen 开始', { actionId: actionId, hostType: hostType, static: USE_STATIC_UI });
    if (USE_STATIC_UI) {
      // 混合：优先用网关 http（避免 file:// 模块/CORS 限制），网关不存活再回退 file:// 离线包
      dbg('静态模式：先探活网关，存活则用 http，失败再用 file://');
      return gatewayAlive().then(function (ok) {
        dbg('静态模式探活结果', ok);
        if (ok) {
          // 网关存活：用 http URL（此时 openOrUpdateTaskPane 需临时走网关）
          var gwUrl = GATEWAY_URL + '/?host=wps&app=' + encodeURIComponent(hostType) + (actionId ? '&action=' + encodeURIComponent(actionId) : '') + '&t=' + Date.now();
          dbg('网关存活，使用 http 任务窗格', gwUrl);
          // 临时关闭静态开关调用原逻辑
          var saved = USE_STATIC_UI;
          USE_STATIC_UI = false;
          try { openOrUpdateTaskPane(actionId); } catch (e) { dbg('http open 失败', e.message||e); }
          USE_STATIC_UI = saved;
          // 若 CreateTaskPane 仍失败，openOrUpdateTaskPane 内部会降级 ShowDialog（http）
          return true;
        } else {
          dbg('网关未存活，回退 file:// 离线包');
          try { openOrUpdateTaskPane(actionId); } catch (e) { dbg('静态 open 失败', e.message||e); }
          return true;
        }
      });
    }
    // 立即预反馈（尝试所有宿主的 ShowDialog）
    try {
      if (BOOTSTRAP_URL.indexOf('http') === 0) {
        var toastUrl = BOOTSTRAP_URL + '?app=' + encodeURIComponent(hostType) + '&toast=1';
        var hs3 = getTaskPaneHosts();
        for (var kk=0; kk<hs3.length; kk++) {
          var hh4 = hs3[kk].obj;
          if (hh4 && typeof hh4.ShowDialog === 'function') {
            dbg('toast ShowDialog @'+hs3[kk].name, toastUrl);
            try { hh4.ShowDialog(toastUrl, 'OfficeAI 正在启动', 420, 180, false); break; } catch(e){ dbg('toast 失败 @'+hs3[kk].name, e.message||e); }
          }
        }
      }
    } catch (e) {}
    return gatewayAlive().then(function (ok) {
      dbg('首次探活结果', ok);
      if (ok) {
        dbg('网关已存活，直接打开 pane');
        try { openOrUpdateTaskPane(actionId); } catch (e) { dbg('open pane 失败', e.message||e); }
        return true;
      }
      if (launching) {
        dbg('已在启动中，提示等待');
        try { var a = getApp(); if (a && a.Alert) a.Alert('OfficeAI 正在启动中，请稍等 2-3 秒后重试…\n\n' + getDebugText().slice(0, 2000)); } catch (e) {}
        return false;
      }
      launching = true;
      dbg('准备拉起网关');
      var shellOk = launchGateway(actionId, hostType);
      if (!shellOk) {
        launching = false;
        dbg('拉起派发失败');
        showLaunchFailed(hostType);
        return false;
      }
      dbg('等待网关就绪轮询');
      return waitForGateway().then(function (started) {
        launching = false;
        dbg('waitForGateway 最终', started);
        if (started) {
          try { openOrUpdateTaskPane(actionId); } catch (e) { dbg('open pane 失败', e.message||e); }
          return true;
        }
        showLaunchFailed(hostType);
        return false;
      });
    }).catch(function (err) {
      launching = false;
      dbg('ensureGatewayThen 异常', err.message||err);
      showLaunchFailed(detectWpsHostType());
      return false;
    });
  }

  function __oa_OnAddinLoad(ribbonUI) {
    __OA_G.OfficeAIRibbonUI = ribbonUI;
    try { window.OfficeAIRibbonUI = ribbonUI; } catch(e){}
    dbg('OnAddinLoad', { GATEWAY_URL: GATEWAY_URL, LAUNCH_COMMAND: LAUNCH_COMMAND });
    gatewayAlive().then(function (ok) { dbg('OnAddinLoad 探活', ok); });
    return true;
  }
  function __oa_GetEnabled(){ return true; }
  function __oa_GetImage(control){
    dbg('GetImage', control && control.Id);
    // WPS 的 GetImage 要求返回插件目录内的相对路径或 file://，http 会显示为 ?
    try {
      // 优先返回相对路径，WPS 会自动解析为插件根
      if (control && control.Id) {
        // 不同按钮可返回不同图标（若存在），否则统一 32
        var iconMap = {
          btnOpenTaskPane: 'assets/icon-32.png',
          btnPolish: 'assets/icon-32.png',
          btnTranslate: 'assets/icon-32.png',
          btnSummarize: 'assets/icon-32.png',
          btnFormula: 'assets/icon-32.png',
          btnPptOutline: 'assets/icon-32.png',
          btnOpenDashboard: 'assets/icon-32.png',
          btnDebugLog: 'assets/icon-32.png'
        };
        if (iconMap[control.Id]) return iconMap[control.Id];
      }
      return 'assets/icon-32.png';
    } catch(e) { return 'assets/icon-32.png'; }
  }
  function __oa_OnAction(control) {
    try {
      var id = control && control.Id ? control.Id : 'btnOpenTaskPane';
      dbg('OnAction 触发', id);
      if (id === 'btnDebugLog') { showDebugLog(); return; }
      if (id === 'btnOpenDashboard') {
        ensureGatewayThen('').then(function (ok) {
          if (ok) try { openDashboard(); } catch (e) { dbg('openDashboard 失败', e.message||e); showLaunchFailed(detectWpsHostType()); }
        }).catch(function () { showLaunchFailed(detectWpsHostType()); });
        return;
      }
      var actionMap = { btnOpenTaskPane: '', btnPolish: 'polish', btnTranslate: 'translate-en', btnSummarize: 'summarize', btnFormula: 'excel-formula', btnPptOutline: 'ppt-outline' };
      ensureGatewayThen(actionMap[id] || '').catch(function () { showLaunchFailed(detectWpsHostType()); });
    } catch (err) {
      dbg('OnAction 异常', err.message||err);
      showLaunchFailed(detectWpsHostType());
    }
  };
  // 暴露给 WPS ribbon（全局函数 + window/globalThis/self 三重挂载，兼容不同 WPS 版本的查找方式）
  try {
    __OA_G.OnAction = __oa_OnAction;
    __OA_G.OnAddinLoad = __oa_OnAddinLoad;
    __OA_G.OnAddInLoad = __oa_OnAddinLoad;
    __OA_G.OnGetEnabled = __oa_GetEnabled;
    __OA_G.GetImage = __oa_GetImage;
    /* placeholder for bootstrap, assigned after */
    try { OnAction = __oa_OnAction; } catch(e){}
    try { OnAddinLoad = __oa_OnAddinLoad; } catch(e){}
    try { OnAddInLoad = __oa_OnAddinLoad; } catch(e){}
    try { OnGetEnabled = __oa_GetEnabled; } catch(e){}
    try { GetImage = __oa_GetImage; } catch(e){}
    try { window.OnAction = __oa_OnAction; } catch(e){}
    try { window.OnAddinLoad = __oa_OnAddinLoad; } catch(e){}
    try { window.OnAddInLoad = __oa_OnAddinLoad; } catch(e){}
    try { window.OnGetEnabled = __oa_GetEnabled; } catch(e){}
    try { window.GetImage = __oa_GetImage; } catch(e){}
    if (typeof self !== 'undefined') {
      try { self.OnAction = __oa_OnAction; } catch(e){}
      try { self.OnAddinLoad = __oa_OnAddinLoad; } catch(e){}
    }
  } catch(e) { dbg('全局挂载失败', e.message||e); }

  __OA_G.OfficeAIBootstrap = window.OfficeAIBootstrap = function (hostParam) {
    var host = hostParam === 'wps' ? detectWpsHostType() : hostParam || 'standalone';
    dbg('OfficeAIBootstrap', host);
    gatewayAlive().then(function (ok) {
      dbg('Bootstrap 探活', ok);
      if (ok) { dbg('Bootstrap 跳转到网关'); window.location.replace(GATEWAY_URL + '/?host=' + encodeURIComponent(host)); return; }
      var shell = document.body || document.documentElement;
      shell.setAttribute('style', 'margin:0;font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;background:#f8fafc;color:#0f172a;display:flex;align-items:center;justify-content:center;min-height:100vh;');
      shell.innerHTML =
        '<div style="max-width:460px;padding:24px">' +
        '<div style="background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:20px">' +
        '<h1 style="font-size:17px;margin:0 0 8px">OfficeAI 服务未启动</h1>' +
        '<p style="font-size:13px;line-height:1.7;color:#475569;margin:0 0 14px">点击下方按钮会在需要时自动拉起 OfficeAI 本地服务，随后自动进入助手。</p>' +
        '<button id="officeai-start" style="width:100%;border:0;border-radius:10px;padding:11px 14px;font-size:14px;font-weight:600;color:#fff;background:#2563eb;cursor:pointer">一键启动 OfficeAI</button>' +
        '<div id="officeai-state" style="font-size:12px;color:#64748b;margin-top:12px;white-space:pre-wrap;max-height:160px;overflow:auto"></div>' +
        '<pre id="officeai-log" style="font-size:10px;background:#0f172a;color:#e2e8f0;padding:8px;border-radius:8px;max-height:140px;overflow:auto;margin-top:10px">' + DEBUG_LOG.slice(-40).join('\n').replace(/</g,'&lt;') + '</pre>' +
        '<p style="font-size:11px;color:#94a3b8;margin-top:12px">启动命令：' + (LAUNCH_COMMAND || LAUNCH_URL || 'OfficeAI.exe') + '<br>网关：' + GATEWAY_URL + '</p>' +
        '</div></div>';
      var stateEl = document.getElementById('officeai-state');
      function setState(t) { if (stateEl) stateEl.textContent = t; dbg('Bootstrap setState', t); }
      var btn = document.getElementById('officeai-start');
      if (btn) btn.onclick = function () {
        dbg('Bootstrap 点击一键启动');
        setState('正在唤起 OfficeAI…');
        launchGateway('', host);
        waitForGateway().then(function (started) {
          dbg('Bootstrap wait 结果', started);
          if (started) window.location.replace(GATEWAY_URL + '/?host=' + encodeURIComponent(host));
          else setState('仍未能启动。请从开始菜单运行「OfficeAI」后重试。' + (LAUNCH_COMMAND ? '（' + LAUNCH_COMMAND + '）' : '') + '\n\n' + DEBUG_LOG.slice(-20).join('\n'));
        });
      };
      setState('');
    });
  };
  } catch (outerErr) {
    try {
      var msg = 'OfficeAI WPS 加载项初始化异常: ' + (outerErr && outerErr.message ? outerErr.message : outerErr) + '\n' + (outerErr && outerErr.stack ? outerErr.stack : '');
      try { console.log(msg); } catch (e) {}
      var __oaErrFn = function(c){ try{ var id=(c&&c.Id)||'unknown'; var t='初始化异常，按钮仍可用: '+id+'\n'+msg.slice(0,1500); if(typeof wps!=='undefined'&&wps.Alert) wps.Alert(t); else if(typeof alert!=='undefined') alert(t); }catch(e){} };
      try { window.OnAction = __oaErrFn; } catch(e){}
      try { __OA_G.OnAction = __oaErrFn; } catch(e){}
      try { OnAction = __oaErrFn; } catch(e){}
      try { window.OnAddinLoad = function(){ return true; }; } catch(e){}
      try { __OA_G.OnAddinLoad = function(){ return true; }; } catch(e){}
    } catch (e2) {}
  }
})();
// 最终兜底校验：若 OnAction 仍未定义，强制定义（兼容 WPS 直接查找全局函数）
try {
  var __OA_F = function(c){ try{ var id=(c&&c.Id)||'unknown'; var t='OfficeAI 兜底: 按钮 '+(id)+' 安装正常但主逻辑未加载，请点“调试日志”查看或重装'; if(typeof wps!=='undefined'&&wps.Alert) wps.Alert(t); else if(typeof alert!=='undefined') alert(t); }catch(e){} };
  var __OA_G3 = (typeof globalThis !== 'undefined' ? globalThis : typeof window !== 'undefined' ? window : typeof self !== 'undefined' ? self : this);
  if (typeof __OA_G3.OnAction !== 'function') __OA_G3.OnAction = __OA_F;
  if (typeof window !== 'undefined' && typeof window.OnAction !== 'function') window.OnAction = __OA_F;
  try { if (typeof OnAction !== 'function') OnAction = __OA_F; } catch(e){}
  if (typeof __OA_G3.OnAddinLoad !== 'function') __OA_G3.OnAddinLoad = function(){ return true; };
  try { if (typeof OnAddinLoad !== 'function') OnAddinLoad = __OA_G3.OnAddinLoad; } catch(e){}
  try { if (typeof OnAddInLoad !== 'function') OnAddInLoad = __OA_G3.OnAddinLoad; } catch(e){}
} catch (e) {}
