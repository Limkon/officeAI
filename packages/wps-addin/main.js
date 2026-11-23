/* global wps, et, wpp, Application */
(function () {
  var GATEWAY_URL = 'http://127.0.0.1:18080';
  var TASKPANE_KEY = 'officeai_taskpane_id';

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

      if (typeof app.CreateTaskPane === 'function') {
        var pane = app.CreateTaskPane(targetUrl, 'OfficeAI 智能助手');
        if (pane) {
          pane.DockPosition = 2; // msoCTPDockPositionRight
          pane.Width = 400;
          pane.Visible = true;
          if (app.PluginStorage && typeof app.PluginStorage.setItem === 'function') {
            app.PluginStorage.setItem(TASKPANE_KEY, pane.ID);
          }
        }
      }
    } catch (err) {
      if (typeof app.ShowDialog === 'function') {
        app.ShowDialog(targetUrl, 'OfficeAI 智能助手', 960, 680, false);
      }
    }
  }

  window.OnAddinLoad = function (ribbonUI) {
    window.OfficeAIRibbonUI = ribbonUI;
    return true;
  };

  window.OnGetEnabled = function () {
    return true;
  };

  window.GetImage = function () {
    return 'assets/icon-32.png';
  };

  window.OnAction = function (control) {
    var id = control && control.Id ? control.Id : 'btnOpenTaskPane';
    if (id === 'btnOpenDashboard') {
      var app = getApp();
      var dashUrl = GATEWAY_URL + '/dashboard';
      if (app && typeof app.ShowDialog === 'function') {
        app.ShowDialog(dashUrl, 'OfficeAI 模型与集成配置中心', 1100, 760, false);
      } else {
        openOrUpdateTaskPane('');
      }
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

    openOrUpdateTaskPane(actionMap[id] || '');
  };
})();
