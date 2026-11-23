import type { OfficeHostApp, OfficeHostSuite } from '@officeai/shared';

/* eslint-disable @typescript-eslint/no-explicit-any */
declare global {
  interface Window {
    Office?: any;
    Word?: any;
    Excel?: any;
    PowerPoint?: any;
    Application?: any;
    wps?: any;
    et?: any;
    wpp?: any;
  }
}

export interface HostContextInfo {
  suite: OfficeHostSuite;
  app: OfficeHostApp;
  label: string;
}

export interface SelectionSnapshot {
  text: string;
  context?: string;
  address?: string;
}

function getWpsApp(): any {
  if (typeof window.Application !== 'undefined' && window.Application) {
    return window.Application;
  }
  if (window.wps?.WpsApplication) return window.wps.WpsApplication();
  if (window.et?.EtApplication) return window.et.EtApplication();
  if (window.wpp?.WppApplication) return window.wpp.WppApplication();
  return null;
}

function matrixToMarkdown(values: unknown[][]): string {
  if (!Array.isArray(values) || values.length === 0) return '';
  const rows = values.slice(0, 50).map((row) =>
    (Array.isArray(row) ? row : [row]).map((cell) =>
      cell === null || cell === undefined ? '' : String(cell).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ')
    )
  );
  if (rows.length === 0) return '';
  const header = `| ${rows[0].join(' | ')} |`;
  const sep = `| ${rows[0].map(() => '---').join(' | ')} |`;
  const body = rows.slice(1).map((r) => `| ${r.join(' | ')} |`);
  return [header, sep, ...body].join('\n');
}

function extractFirstFormula(text: string): string | null {
  const codeMatch = text.match(/```[a-zA-Z]*\n?(=[^\n`]+)/);
  if (codeMatch) return codeMatch[1].trim();
  const lineMatch = text.match(/^(=[A-Z0-9_.]+\(.*\))/m);
  if (lineMatch) return lineMatch[1].trim();
  return null;
}

export function getHostLabel(app: OfficeHostApp): string {
  switch (app) {
    case 'word':
      return 'Microsoft Word';
    case 'excel':
      return 'Microsoft Excel';
    case 'powerpoint':
      return 'Microsoft PowerPoint';
    case 'wps-writer':
      return 'WPS 文字';
    case 'wps-spreadsheets':
      return 'WPS 表格';
    case 'wps-presentation':
      return 'WPS 演示';
    default:
      return '独立桌面/预览模式';
  }
}

export async function detectHostContext(): Promise<HostContextInfo> {
  const params = new URLSearchParams(window.location.search);
  const hostHint = params.get('host');
  const appHint = params.get('app') as OfficeHostApp | null;

  // 1. Check Microsoft Office.js
  if (typeof window.Office !== 'undefined' && window.Office.onReady) {
    try {
      const info = await new Promise<{ host?: unknown }>((resolve) => {
        const timer = setTimeout(() => resolve({}), 1200);
        window.Office.onReady((res: { host?: unknown }) => {
          clearTimeout(timer);
          resolve(res || {});
        });
      });

      const host = info?.host;
      if (host === window.Office.HostType?.Word) {
        return { suite: 'ms-office', app: 'word', label: getHostLabel('word') };
      }
      if (host === window.Office.HostType?.Excel) {
        return { suite: 'ms-office', app: 'excel', label: getHostLabel('excel') };
      }
      if (host === window.Office.HostType?.PowerPoint) {
        return { suite: 'ms-office', app: 'powerpoint', label: getHostLabel('powerpoint') };
      }
    } catch {
      // fall through
    }
  }

  // 2. Check WPS JSAPI
  const wpsApp = getWpsApp();
  if (wpsApp || hostHint === 'wps') {
    if (appHint && appHint.startsWith('wps-')) {
      return { suite: 'wps', app: appHint, label: getHostLabel(appHint) };
    }
    try {
      if (wpsApp?.ActiveWorkbook || String(wpsApp?.Name || '').toLowerCase().includes('et')) {
        return { suite: 'wps', app: 'wps-spreadsheets', label: getHostLabel('wps-spreadsheets') };
      }
      if (wpsApp?.ActivePresentation || String(wpsApp?.Name || '').toLowerCase().includes('wpp')) {
        return { suite: 'wps', app: 'wps-presentation', label: getHostLabel('wps-presentation') };
      }
    } catch {
      // default writer
    }
    return { suite: 'wps', app: 'wps-writer', label: getHostLabel('wps-writer') };
  }

  if (appHint && appHint !== 'standalone') {
    return {
      suite: appHint.startsWith('wps-') ? 'wps' : 'ms-office',
      app: appHint,
      label: getHostLabel(appHint)
    };
  }

  return { suite: 'standalone', app: 'standalone', label: getHostLabel('standalone') };
}

export async function readHostSelection(app: OfficeHostApp): Promise<SelectionSnapshot> {
  // --- Microsoft Word ---
  if (app === 'word' && window.Word?.run) {
    return window.Word.run(async (context: any) => {
      const sel = context.document.getSelection();
      sel.load('text');
      await context.sync();
      const text = (sel.text || '').trim();
      if (text) return { text };
      const body = context.document.body;
      body.load('text');
      await context.sync();
      return { text: '', context: (body.text || '').slice(0, 4000) };
    });
  }

  // --- Microsoft Excel ---
  if (app === 'excel' && window.Excel?.run) {
    return window.Excel.run(async (context: any) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values', 'formulas']);
      await context.sync();
      const md = matrixToMarkdown(range.values || []);
      return {
        text: md,
        context: `选区地址: ${range.address}\n${md}`,
        address: range.address
      };
    });
  }

  // --- Microsoft PowerPoint ---
  if (app === 'powerpoint' && window.Office?.context?.document) {
    return new Promise((resolve) => {
      window.Office.context.document.getSelectedDataAsync(
        window.Office.CoercionType.Text,
        (result: any) => {
          if (result.status === window.Office.AsyncResultStatus.Succeeded) {
            resolve({ text: String(result.value || '').trim() });
          } else {
            resolve({ text: '' });
          }
        }
      );
    });
  }

  // --- WPS Writer ---
  const wpsApp = getWpsApp();
  if (app === 'wps-writer' && wpsApp) {
    try {
      const selText = String(wpsApp.Selection?.Text || '').trim();
      if (selText) return { text: selText };
      const docText = String(wpsApp.ActiveDocument?.Content?.Text || '').slice(0, 4000);
      return { text: '', context: docText };
    } catch {
      return { text: '' };
    }
  }

  // --- WPS Spreadsheets ---
  if (app === 'wps-spreadsheets' && wpsApp) {
    try {
      const sel = wpsApp.Selection;
      const address = sel?.Address ? String(sel.Address()) : '';
      const rawVal = sel?.Value2;
      let md = '';
      if (Array.isArray(rawVal)) {
        md = matrixToMarkdown(rawVal);
      } else if (rawVal !== undefined && rawVal !== null) {
        md = String(rawVal);
      }
      return {
        text: md,
        context: address ? `选区地址: ${address}\n${md}` : md,
        address
      };
    } catch {
      return { text: '' };
    }
  }

  // --- WPS Presentation ---
  if (app === 'wps-presentation' && wpsApp) {
    try {
      const selText = String(wpsApp.ActiveWindow?.Selection?.TextRange?.Text || '').trim();
      return { text: selText };
    } catch {
      return { text: '' };
    }
  }

  return { text: '' };
}

export async function writeHostContent(
  app: OfficeHostApp,
  text: string,
  mode: 'replace' | 'insert' | 'formula'
): Promise<{ ok: boolean; message: string }> {
  const cleanText = text.trim();
  if (!cleanText) {
    return { ok: false, message: '没有可写入的内容' };
  }

  // --- Microsoft Word ---
  if (app === 'word' && window.Word?.run) {
    await window.Word.run(async (context: any) => {
      const sel = context.document.getSelection();
      const loc = mode === 'replace' ? window.Word.InsertLocation.replace : window.Word.InsertLocation.after;
      sel.insertText(mode === 'insert' ? `\n${cleanText}` : cleanText, loc);
      await context.sync();
    });
    return { ok: true, message: mode === 'replace' ? '已替换 Word 选中内容' : '已插入到 Word 光标后' };
  }

  // --- Microsoft Excel ---
  if (app === 'excel' && window.Excel?.run) {
    const formula = mode === 'formula' ? extractFirstFormula(cleanText) || cleanText : null;
    await window.Excel.run(async (context: any) => {
      const range = context.workbook.getSelectedRange();
      const cell = range.getCell(0, 0);
      if (formula && formula.startsWith('=')) {
        cell.formulas = [[formula]];
      } else {
        cell.values = [[cleanText]];
      }
      await context.sync();
    });
    return { ok: true, message: formula ? `已写入公式 ${formula}` : '已写入 Excel 单元格' };
  }

  // --- Microsoft PowerPoint ---
  if (app === 'powerpoint' && window.Office?.context?.document) {
    await new Promise<void>((resolve, reject) => {
      window.Office.context.document.setSelectedDataAsync(
        cleanText,
        { coercionType: window.Office.CoercionType.Text },
        (res: any) => {
          if (res.status === window.Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(res.error?.message || 'PowerPoint 写入失败'));
        }
      );
    });
    return { ok: true, message: '已写入 PowerPoint 文本框' };
  }

  // --- WPS Writer ---
  const wpsApp = getWpsApp();
  if (app === 'wps-writer' && wpsApp) {
    const sel = wpsApp.Selection;
    if (!sel) return { ok: false, message: '未获取到 WPS 光标选区' };
    if (mode === 'insert') {
      sel.Collapse(0); // wdCollapseEnd
      sel.TypeParagraph();
    }
    sel.TypeText(cleanText);
    return { ok: true, message: mode === 'replace' ? '已替换 WPS 文字选区' : '已插入到 WPS 文档' };
  }

  // --- WPS Spreadsheets ---
  if (app === 'wps-spreadsheets' && wpsApp) {
    const cell = wpsApp.ActiveCell || wpsApp.Selection;
    if (!cell) return { ok: false, message: '未获取到 WPS 表格单元格' };
    const formula = mode === 'formula' ? extractFirstFormula(cleanText) || cleanText : null;
    if (formula && formula.startsWith('=')) {
      cell.Formula = formula;
      return { ok: true, message: `已写入 WPS 表格公式 ${formula}` };
    }
    cell.Value2 = cleanText;
    return { ok: true, message: '已写入 WPS 表格单元格' };
  }

  // --- WPS Presentation ---
  if (app === 'wps-presentation' && wpsApp) {
    const tr = wpsApp.ActiveWindow?.Selection?.TextRange;
    if (tr) {
      tr.Text = mode === 'insert' ? `${tr.Text || ''}\n${cleanText}` : cleanText;
      return { ok: true, message: '已写入 WPS 演示文本框' };
    }
    const slide = wpsApp.ActiveWindow?.View?.Slide;
    if (slide?.NotesPage?.Shapes?.Placeholders) {
      slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = cleanText;
      return { ok: true, message: '已写入当前幻灯片演讲备注' };
    }
    return { ok: false, message: '请先在幻灯片中选中一个文本框' };
  }

  return { ok: false, message: '当前处于独立预览模式，已为你保留结果，可点击复制按钮使用。' };
}
