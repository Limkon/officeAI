import type { OfficeHostApp } from './types.js';

export type OfficeActionType =
  | 'insert_text'
  | 'replace_selection'
  | 'find_replace'
  | 'insert_table'
  | 'set_cells'
  | 'format_range'
  | 'create_sheet'
  | 'sort_range'
  | 'add_slide'
  | 'set_speaker_notes';

export interface OfficeActionPayload {
  action: OfficeActionType;
  summary: string;
  // Word / WPS Writer
  text?: string;
  position?: 'cursor' | 'start' | 'end';
  find?: string;
  replaceWith?: string;
  matchCase?: boolean;
  // Tables / Excel / WPS Spreadsheets
  range?: string; // e.g. "A1:C3" or "E2" or "active"
  sheetName?: string;
  values?: (string | number | boolean)[][];
  formula?: string;
  bold?: boolean;
  fillColor?: string; // hex like "#EEF2FF"
  fontColor?: string; // hex like "#1E293B"
  numberFormat?: string; // e.g. "0.00%" or "#,##0.00"
  autoFit?: boolean;
  sortColumnIndex?: number; // 0-based
  ascending?: boolean;
  // PPT / WPS Presentation
  title?: string;
  bullets?: string[];
  notes?: string;
}

export interface ParsedAssistantActions {
  displayText: string;
  actions: OfficeActionPayload[];
}

const VALID_ACTIONS = new Set<string>([
  'insert_text',
  'replace_selection',
  'find_replace',
  'insert_table',
  'set_cells',
  'format_range',
  'create_sheet',
  'sort_range',
  'add_slide',
  'set_speaker_notes'
]);

export function buildOfficeAgentSystemPrompt(
  basePrompt: string,
  hostApp: OfficeHostApp
): string {
  const hostNote = (() => {
    switch (hostApp) {
      case 'word':
      case 'wps-writer':
        return [
          '当前宿主为文字文档（Word / WPS 文字），可执行的操作 action 类型包括：',
          '1. "replace_selection": { "action": "replace_selection", "summary": "替换选中文本", "text": "..." }',
          '2. "insert_text": { "action": "insert_text", "summary": "插入文本", "text": "...", "position": "cursor|end|start" }',
          '3. "find_replace": { "action": "find_replace", "summary": "查找替换", "find": "旧词", "replaceWith": "新词", "matchCase": false }',
          '4. "insert_table": { "action": "insert_table", "summary": "插入表格", "values": [["列1","列2"],["a","b"]], "position": "cursor|end" }'
        ].join('\n');
      case 'excel':
      case 'wps-spreadsheets':
        return [
          '当前宿主为电子表格（Excel / WPS 表格），可执行的操作 action 类型包括：',
          '1. "set_cells": { "action": "set_cells", "summary": "写入单元格/公式/表格", "range": "A1:C3 或 E2 或 active", "sheetName": "可选工作表名", "values": [["标题1","标题2"],["数据", "=SUM(B2:B10)"]] } （如果只写单个公式也可用 "formula": "=SUM(B2:B10)", "range": "E2"）',
          '2. "format_range": { "action": "format_range", "summary": "设置格式", "range": "A1:C1", "bold": true, "fillColor": "#E0E7FF", "fontColor": "#1E293B", "numberFormat": "0.00%", "autoFit": true }',
          '3. "create_sheet": { "action": "create_sheet", "summary": "新建工作表", "sheetName": "汇总", "values": [["列1","列2"]] }',
          '4. "sort_range": { "action": "sort_range", "summary": "排序区域", "range": "A1:D10", "sortColumnIndex": 1, "ascending": false }'
        ].join('\n');
      case 'powerpoint':
      case 'wps-presentation':
        return [
          '当前宿主为演示文稿（PowerPoint / WPS 演示），可执行的操作 action 类型包括：',
          '1. "replace_selection": { "action": "replace_selection", "summary": "修改当前文本框", "text": "..." }',
          '2. "add_slide": { "action": "add_slide", "summary": "新增幻灯片", "title": "页面标题", "bullets": ["要点1", "要点2"] }',
          '3. "set_speaker_notes": { "action": "set_speaker_notes", "summary": "写入演讲备注", "notes": "..." }'
        ].join('\n');
      default:
        return '当前为独立预览模式。若用户要求生成 Word/Excel/PPT 操作，仍可返回 <office_action> 演示结构。';
    }
  })();

  return [
    basePrompt,
    '',
    '【文档直接操作能力】',
    '当用户的指令明确要求修改、写入、生成表格/公式、查找替换、排版或操作当前文档时，除了简短文字说明外，请在回复末尾附带一个或多个 <office_action>JSON</office_action> 标签来执行实际操作；如果用户只是普通问答/咨询，则无需输出 <office_action> 标签。',
    hostNote,
    '注意：每个 <office_action>...</office_action> 内部必须是合法的单个 JSON 对象，不要包含注释或多余文本。'
  ].join('\n');
}

/**
 * Removes complete and trailing partial <office_action> tags while streaming
 * so raw JSON never flickers in the chat bubble.
 */
export function stripStreamingOfficeActionTags(rawText: string): string {
  // Remove complete blocks
  let cleaned = rawText.replace(/<office_action>\s*[\s\S]*?\s*<\/office_action>/gi, '');
  // Remove incomplete opening/body at tail: e.g. "<office_action>..." or "<office_"
  cleaned = cleaned.replace(/<office_action[\s\S]*$/i, '');
  cleaned = cleaned.replace(/<(?:o(?:f(?:f(?:i(?:c(?:e(?:_(?:a(?:c(?:t(?:i(?:o(?:n)?)?)?)?)?)?)?)?)?)?)?)?)?$/i, '');
  return cleaned.trim();
}

export function parseOfficeActions(rawText: string): ParsedAssistantActions {
  const actions: OfficeActionPayload[] = [];
  const regex = /<office_action>\s*([\s\S]*?)\s*<\/office_action>/gi;

  const displayText = rawText
    .replace(regex, (fullMatch, jsonBody: string) => {
      const cleaned = jsonBody
        .trim()
        .replace(/^```(?:json)?\s*/i, '')
        .replace(/\s*```$/i, '')
        .trim();
      try {
        const parsed = JSON.parse(cleaned) as OfficeActionPayload | OfficeActionPayload[];
        const items = Array.isArray(parsed) ? parsed : [parsed];
        let extracted = false;
        for (const item of items) {
          if (item && typeof item.action === 'string' && VALID_ACTIONS.has(item.action)) {
            actions.push({
              ...item,
              summary: typeof item.summary === 'string' && item.summary.trim() ? item.summary : item.action
            });
            extracted = true;
          }
        }
        return extracted ? '' : fullMatch;
      } catch {
        return fullMatch;
      }
    })
    .trim();

  return { displayText, actions };
}
