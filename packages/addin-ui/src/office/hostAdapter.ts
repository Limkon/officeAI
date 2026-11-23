import type {
  OfficeActionPayload,
  OfficeHostApp,
  OfficeHostSuite
} from '@officeai/shared';

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

export interface ActionExecutionResult {
  ok: boolean;
  message: string;
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

function getWpsAddress(sel: any): string {
  if (!sel) return '';
  try {
    if (typeof sel.Address === 'function') return String(sel.Address() || '');
    if (typeof sel.Address === 'string') return sel.Address;
  } catch {
    // ignore
  }
  return '';
}

function sanitizeSheetName(raw: string): string {
  const cleaned = raw.replace(/[\\/?*[\]:]/g, '_').trim();
  const base = cleaned || `AI_${Date.now().toString().slice(-4)}`;
  return base.slice(0, 31);
}

function matrixToMarkdown(values: unknown[][]): string {
  if (!Array.isArray(values) || values.length === 0) return '';
  const rows = values.slice(0, 50).map((row) =>
    (Array.isArray(row) ? row : [row]).map((cell) =>
      cell === null || cell === undefined
        ? ''
        : String(cell).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ')
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

function hexToRgbOleColor(hex?: string): number | null {
  if (!hex) return null;
  const clean = hex.replace('#', '').trim();
  if (!/^[0-9a-fA-F]{6}$/.test(clean)) return null;
  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);
  // Excel/WPS COM Color uses BGR integer
  return r + g * 256 + b * 65536;
}

function normalizeRectangularValues(
  values: (string | number | boolean)[][]
): (string | number | boolean)[][] {
  if (values.length === 0) return [['']];
  const maxCols = Math.max(...values.map((r) => (Array.isArray(r) ? r.length : 1)), 1);
  return values.map((row) => {
    const arr = Array.isArray(row) ? [...row] : [row];
    while (arr.length < maxCols) arr.push('');
    return arr;
  });
}

function writeWpsMatrix(
  sheet: any,
  startRow: number,
  startCol: number,
  matrix: (string | number | boolean)[][]
): void {
  const rows = matrix.length;
  const cols = matrix[0].length;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const val = matrix[r][c];
      const cell = sheet.Cells(startRow + r, startCol + c);
      if (typeof val === 'string' && val.startsWith('=')) {
        cell.Formula = val;
      } else {
        cell.Value2 = val ?? '';
      }
    }
  }
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
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values', 'formulas']);
      await context.sync();
      const md = matrixToMarkdown(range.values || []);
      return {
        text: md,
        context: `当前工作表: ${sheet.name}\n选区地址: ${range.address}\n${md}`,
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
      const sheetName = String(wpsApp.ActiveSheet?.Name || '');
      const sel = wpsApp.Selection;
      const address = getWpsAddress(sel);
      const rawVal = sel?.Value2;
      let md = '';
      if (Array.isArray(rawVal)) {
        md = matrixToMarkdown(rawVal);
      } else if (rawVal !== undefined && rawVal !== null) {
        md = String(rawVal);
      }
      return {
        text: md,
        context: `${sheetName ? `当前工作表: ${sheetName}\n` : ''}${address ? `选区地址: ${address}\n` : ''}${md}`,
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
): Promise<ActionExecutionResult> {
  const cleanText = text.trim();
  if (!cleanText) {
    return { ok: false, message: '没有可写入的内容' };
  }

  if (mode === 'formula') {
    const formula = extractFirstFormula(cleanText) || cleanText;
    return executeOfficeAction(app, {
      action: 'set_cells',
      summary: '写入公式/单元格',
      range: 'active',
      ...(formula.startsWith('=') ? { formula } : { values: [[cleanText]] })
    });
  }

  if (mode === 'replace') {
    return executeOfficeAction(app, {
      action: 'replace_selection',
      summary: '替换选区',
      text: cleanText
    });
  }

  return executeOfficeAction(app, {
    action: 'insert_text',
    summary: '插入文本',
    text: cleanText,
    position: 'cursor'
  });
}

export async function executeOfficeAction(
  app: OfficeHostApp,
  payload: OfficeActionPayload
): Promise<ActionExecutionResult> {
  if (app === 'standalone') {
    return {
      ok: false,
      message: '当前为独立浏览器/预览模式，请在 Word / Excel / WPS 中打开右侧边栏以直接执行文档操作。'
    };
  }

  // ==========================================
  // 1. MICROSOFT WORD
  // ==========================================
  if (app === 'word' && window.Word?.run) {
    return window.Word.run(async (context: any) => {
      switch (payload.action) {
        case 'replace_selection': {
          const sel = context.document.getSelection();
          sel.insertText(payload.text || '', window.Word.InsertLocation.replace);
          await context.sync();
          return { ok: true, message: '已替换 Word 选中内容' };
        }
        case 'insert_text': {
          const text = payload.text || '';
          if (payload.position === 'end') {
            context.document.body.insertText(`\n${text}`, window.Word.InsertLocation.end);
          } else if (payload.position === 'start') {
            context.document.body.insertText(`${text}\n`, window.Word.InsertLocation.start);
          } else {
            const sel = context.document.getSelection();
            sel.insertText(`\n${text}`, window.Word.InsertLocation.after);
          }
          await context.sync();
          return { ok: true, message: '已向 Word 文档写入内容' };
        }
        case 'find_replace': {
          if (!payload.find) return { ok: false, message: '缺少要查找的文本 (find)' };
          if (payload.find.length > 255) {
            return { ok: false, message: 'Word 查找关键词不能超过 255 个字符' };
          }
          const results = context.document.body.search(payload.find, {
            matchCase: Boolean(payload.matchCase)
          });
          results.load('items');
          await context.sync();
          for (const item of results.items) {
            item.insertText(payload.replaceWith ?? '', window.Word.InsertLocation.replace);
          }
          await context.sync();
          return {
            ok: true,
            message: `Word 查找替换完成，共替换 ${results.items.length} 处`
          };
        }
        case 'insert_table': {
          const matrix = normalizeRectangularValues(payload.values || [['列1', '列2']]).map((r) =>
            r.map((c) => String(c ?? ''))
          );
          const rows = matrix.length;
          const cols = matrix[0].length;
          if (payload.position === 'end') {
            context.document.body.insertTable(rows, cols, window.Word.InsertLocation.end, matrix);
          } else {
            const sel = context.document.getSelection();
            sel.insertTable(rows, cols, window.Word.InsertLocation.after, matrix);
          }
          await context.sync();
          return { ok: true, message: `已在 Word 插入 ${rows}×${cols} 表格` };
        }
        default:
          return { ok: false, message: `Word 暂不支持操作类型: ${payload.action}` };
      }
    });
  }

  // ==========================================
  // 2. MICROSOFT EXCEL
  // ==========================================
  if (app === 'excel' && window.Excel?.run) {
    return window.Excel.run(async (context: any) => {
      const getTargetSheet = () =>
        payload.sheetName
          ? context.workbook.worksheets.getItem(payload.sheetName)
          : context.workbook.worksheets.getActiveWorksheet();

      const getTargetRange = (sheet: any, addr?: string) => {
        if (!addr || addr.toLowerCase() === 'active') {
          return context.workbook.getSelectedRange();
        }
        return sheet.getRange(addr);
      };

      switch (payload.action) {
        case 'set_cells': {
          const sheet = getTargetSheet();
          const baseRange = getTargetRange(sheet, payload.range);
          if (payload.formula) {
            const cell = baseRange.getCell(0, 0);
            cell.formulas = [[payload.formula]];
            await context.sync();
            return { ok: true, message: `已写入 Excel 公式: ${payload.formula}` };
          }
          const matrix = normalizeRectangularValues(payload.values || [['']]);
          const rows = matrix.length;
          const cols = matrix[0].length;
          const resized = baseRange.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
          const hasFormula = matrix.some((r) =>
            r.some((v) => typeof v === 'string' && v.startsWith('='))
          );
          if (hasFormula) {
            resized.formulas = matrix;
          } else {
            resized.values = matrix;
          }
          await context.sync();
          return { ok: true, message: `已写入 Excel 区域 (${rows}×${cols})` };
        }
        case 'format_range': {
          const sheet = getTargetSheet();
          const range = getTargetRange(sheet, payload.range);
          if (payload.bold !== undefined) range.format.font.bold = payload.bold;
          if (payload.fillColor) range.format.fill.color = payload.fillColor;
          if (payload.fontColor) range.format.font.color = payload.fontColor;
          if (payload.numberFormat) {
            range.load(['rowCount', 'columnCount']);
            await context.sync();
            const fmtMatrix = Array.from({ length: range.rowCount }, () =>
              Array.from({ length: range.columnCount }, () => payload.numberFormat)
            );
            range.numberFormat = fmtMatrix;
          }
          if (payload.autoFit) {
            range.format.autofitColumns();
          }
          await context.sync();
          return { ok: true, message: '已应用 Excel 单元格样式' };
        }
        case 'create_sheet': {
          const baseName = sanitizeSheetName(payload.sheetName || `AI_${Date.now().toString().slice(-4)}`);
          const existing = context.workbook.worksheets.getItemOrNullObject(baseName);
          existing.load('isNullObject');
          await context.sync();
          const finalName = existing.isNullObject
            ? baseName
            : sanitizeSheetName(`${baseName.slice(0, 26)}_${Date.now().toString().slice(-4)}`);
          const sheet = context.workbook.worksheets.add(finalName);
          sheet.activate();
          if (payload.values && payload.values.length > 0) {
            const matrix = normalizeRectangularValues(payload.values);
            const target = sheet
              .getRange('A1')
              .getResizedRange(matrix.length - 1, matrix[0].length - 1);
            const hasFormula = matrix.some((r) =>
              r.some((v) => typeof v === 'string' && v.startsWith('='))
            );
            if (hasFormula) {
              target.formulas = matrix;
            } else {
              target.values = matrix;
            }
          }
          await context.sync();
          return { ok: true, message: `已新建并切换到工作表「${finalName}」` };
        }
        case 'sort_range': {
          const sheet = getTargetSheet();
          const range = getTargetRange(sheet, payload.range);
          range.sort.apply([
            {
              key: payload.sortColumnIndex ?? 0,
              ascending: payload.ascending ?? true
            }
          ]);
          await context.sync();
          return { ok: true, message: 'Excel 区域排序完成' };
        }
        default:
          return { ok: false, message: `Excel 暂不支持操作类型: ${payload.action}` };
      }
    });
  }

  // ==========================================
  // 3. MICROSOFT POWERPOINT
  // ==========================================
  if (app === 'powerpoint') {
    if (payload.action === 'replace_selection' && window.Office?.context?.document) {
      await new Promise<void>((resolve, reject) => {
        window.Office.context.document.setSelectedDataAsync(
          payload.text || '',
          { coercionType: window.Office.CoercionType.Text },
          (res: any) => {
            if (res.status === window.Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(res.error?.message || '写入失败'));
          }
        );
      });
      return { ok: true, message: '已更新 PowerPoint 选中文本框' };
    }
    if (payload.action === 'add_slide' && window.PowerPoint?.run) {
      return window.PowerPoint.run(async (context: any) => {
        const slides = context.presentation.slides;
        slides.add();
        slides.load('items');
        await context.sync();
        const newest = slides.items[slides.items.length - 1];
        if (newest?.shapes?.addTextBox) {
          if (payload.title) {
            newest.shapes.addTextBox(payload.title, {
              left: 48,
              top: 36,
              width: 620,
              height: 60
            });
          }
          if (payload.bullets && payload.bullets.length > 0) {
            newest.shapes.addTextBox(payload.bullets.map((b) => `• ${b}`).join('\n'), {
              left: 48,
              top: 112,
              width: 620,
              height: 360
            });
          }
          await context.sync();
        }
        return { ok: true, message: `已新增 PowerPoint 幻灯片「${payload.title || '新页面'}」` };
      });
    }
    return { ok: false, message: `PowerPoint 暂不支持此操作: ${payload.action}` };
  }

  // ==========================================
  // 4. WPS WRITER (文字)
  // ==========================================
  const wpsApp = getWpsApp();
  if (app === 'wps-writer' && wpsApp) {
    const doc = wpsApp.ActiveDocument;
    const sel = wpsApp.Selection;
    switch (payload.action) {
      case 'replace_selection': {
        if (!sel) return { ok: false, message: '未获取到 WPS 光标选区' };
        sel.TypeText(payload.text || '');
        return { ok: true, message: '已替换 WPS 文字选区' };
      }
      case 'insert_text': {
        const text = payload.text || '';
        if (payload.position === 'end' && doc?.Content) {
          doc.Content.InsertAfter(`\n${text}`);
        } else if (payload.position === 'start' && doc?.Content) {
          doc.Content.InsertBefore(`${text}\n`);
        } else if (sel) {
          sel.Collapse(0);
          sel.TypeParagraph();
          sel.TypeText(text);
        }
        return { ok: true, message: '已向 WPS 文档插入内容' };
      }
      case 'find_replace': {
        if (!payload.find || !doc?.Content?.Find) {
          return { ok: false, message: '无法访问 WPS 查找替换接口' };
        }
        const find = doc.Content.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = payload.find;
        find.Replacement.Text = payload.replaceWith ?? '';
        find.MatchCase = Boolean(payload.matchCase);
        find.Execute(
          payload.find,
          Boolean(payload.matchCase),
          false,
          false,
          false,
          false,
          true,
          1,
          false,
          payload.replaceWith ?? '',
          2
        );
        return { ok: true, message: `已完成 WPS 全文替换「${payload.find}」→「${payload.replaceWith ?? ''}」` };
      }
      case 'insert_table': {
        if (!doc?.Tables || !sel) return { ok: false, message: '未获取到 WPS 文档对象' };
        const matrix = normalizeRectangularValues(payload.values || [['列1', '列2']]);
        const rows = matrix.length;
        const cols = matrix[0].length;
        const endPos = Math.max(0, (doc.Content?.End ?? 1) - 1);
        const anchorRange = payload.position === 'end' ? doc.Range(endPos, endPos) : sel.Range;
        const tbl = doc.Tables.Add(anchorRange, rows, cols);
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            tbl.Cell(r + 1, c + 1).Range.Text = String(matrix[r][c] ?? '');
          }
        }
        return { ok: true, message: `已在 WPS 文字插入 ${rows}×${cols} 表格` };
      }
      default:
        return { ok: false, message: `WPS 文字暂不支持操作: ${payload.action}` };
    }
  }

  // ==========================================
  // 5. WPS SPREADSHEETS (表格)
  // ==========================================
  if (app === 'wps-spreadsheets' && wpsApp) {
    const wb = wpsApp.ActiveWorkbook;
    const getSheet = () =>
      payload.sheetName && wb?.Sheets ? wb.Sheets.Item(payload.sheetName) : wpsApp.ActiveSheet;
    const getStartCell = (sheet: any) => {
      if (!payload.range || payload.range.toLowerCase() === 'active') {
        return wpsApp.ActiveCell || wpsApp.Selection;
      }
      return sheet.Range(payload.range);
    };

    switch (payload.action) {
      case 'set_cells': {
        const sheet = getSheet();
        const target = getStartCell(sheet);
        if (payload.formula) {
          target.Cells(1, 1).Formula = payload.formula;
          return { ok: true, message: `已写入 WPS 表格公式: ${payload.formula}` };
        }
        const matrix = normalizeRectangularValues(payload.values || [['']]);
        const rows = matrix.length;
        const cols = matrix[0].length;
        writeWpsMatrix(sheet, target.Row, target.Column, matrix);
        return { ok: true, message: `已写入 WPS 表格区域 (${rows}×${cols})` };
      }
      case 'format_range': {
        const sheet = getSheet();
        const range = getStartCell(sheet);
        if (payload.bold !== undefined) range.Font.Bold = payload.bold;
        const bgOle = hexToRgbOleColor(payload.fillColor);
        if (bgOle !== null) range.Interior.Color = bgOle;
        const fgOle = hexToRgbOleColor(payload.fontColor);
        if (fgOle !== null) range.Font.Color = fgOle;
        if (payload.numberFormat) range.NumberFormat = payload.numberFormat;
        if (payload.autoFit) range.Columns.AutoFit();
        return { ok: true, message: '已应用 WPS 表格样式' };
      }
      case 'create_sheet': {
        if (!wb?.Sheets) return { ok: false, message: '未检测到活动工作簿' };
        const newSheet = wb.Sheets.Add();
        const desired = sanitizeSheetName(payload.sheetName || `AI_${Date.now().toString().slice(-4)}`);
        try {
          newSheet.Name = desired;
        } catch {
          newSheet.Name = sanitizeSheetName(`${desired.slice(0, 26)}_${Date.now().toString().slice(-4)}`);
        }
        if (payload.values && payload.values.length > 0) {
          const matrix = normalizeRectangularValues(payload.values);
          writeWpsMatrix(newSheet, 1, 1, matrix);
        }
        return { ok: true, message: `已新建 WPS 工作表「${newSheet.Name}」` };
      }
      case 'sort_range': {
        const sheet = getSheet();
        const range = getStartCell(sheet);
        const colIdx = (payload.sortColumnIndex ?? 0) + 1;
        const keyCell = range.Cells(1, colIdx);
        range.Sort(keyCell, payload.ascending === false ? 2 : 1, undefined, undefined, undefined, undefined, undefined, 1);
        return { ok: true, message: 'WPS 表格排序完成' };
      }
      default:
        return { ok: false, message: `WPS 表格暂不支持操作: ${payload.action}` };
    }
  }

  // ==========================================
  // 6. WPS PRESENTATION (演示)
  // ==========================================
  if (app === 'wps-presentation' && wpsApp) {
    const pres = wpsApp.ActivePresentation;
    switch (payload.action) {
      case 'replace_selection': {
        const tr = wpsApp.ActiveWindow?.Selection?.TextRange;
        if (!tr) return { ok: false, message: '请先选中 WPS 演示中的文本框' };
        tr.Text = payload.text || '';
        return { ok: true, message: '已更新 WPS 演示文本框' };
      }
      case 'add_slide': {
        if (!pres?.Slides) return { ok: false, message: '未获取到活动演示文稿' };
        const slide = pres.Slides.Add(pres.Slides.Count + 1, 2);
        if (payload.title && slide.Shapes?.Title?.TextFrame?.TextRange) {
          slide.Shapes.Title.TextFrame.TextRange.Text = payload.title;
        }
        if (payload.bullets && payload.bullets.length > 0 && slide.Shapes?.Placeholders?.Item(2)) {
          slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = payload.bullets.join('\n');
        }
        return { ok: true, message: `已新增幻灯片「${payload.title || '新页面'}」` };
      }
      case 'set_speaker_notes': {
        const slide = wpsApp.ActiveWindow?.View?.Slide;
        if (!slide?.NotesPage?.Shapes?.Placeholders) {
          return { ok: false, message: '无法访问当前幻灯片演讲备注' };
        }
        slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = payload.notes || payload.text || '';
        return { ok: true, message: '已写入当前幻灯片演讲备注' };
      }
      default:
        return { ok: false, message: `WPS 演示暂不支持操作: ${payload.action}` };
    }
  }

  return { ok: false, message: '未检测到可操作的 Office / WPS 宿主环境' };
}
