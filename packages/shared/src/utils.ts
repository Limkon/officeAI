/**
 * Incrementally extracts <think>...</think> blocks from streaming text
 * (used by DeepSeek-R1, QwQ, GLM-Z1, Ollama reasoning models).
 */
export class ThinkTagSplitter {
  private buffer = '';
  private inThink = false;

  push(chunk: string): { content: string; reasoning: string } {
    this.buffer += chunk;
    let content = '';
    let reasoning = '';

    while (this.buffer.length > 0) {
      if (!this.inThink) {
        const openIdx = this.buffer.indexOf('<think>');
        if (openIdx === -1) {
          // Keep up to 6 trailing chars in case '<think>' spans chunks
          const safeLen = Math.max(0, this.buffer.length - 6);
          if (safeLen === 0) break;
          content += this.buffer.slice(0, safeLen);
          this.buffer = this.buffer.slice(safeLen);
          break;
        }
        content += this.buffer.slice(0, openIdx);
        this.buffer = this.buffer.slice(openIdx + '<think>'.length);
        this.inThink = true;
      } else {
        const closeIdx = this.buffer.indexOf('</think>');
        if (closeIdx === -1) {
          const safeLen = Math.max(0, this.buffer.length - 7);
          if (safeLen === 0) break;
          reasoning += this.buffer.slice(0, safeLen);
          this.buffer = this.buffer.slice(safeLen);
          break;
        }
        reasoning += this.buffer.slice(0, closeIdx);
        this.buffer = this.buffer.slice(closeIdx + '</think>'.length);
        this.inThink = false;
      }
    }

    return { content, reasoning };
  }

  flush(): { content: string; reasoning: string } {
    const rest = this.buffer;
    this.buffer = '';
    if (this.inThink) {
      return { content: '', reasoning: rest };
    }
    return { content: rest, reasoning: '' };
  }
}

export function maskApiKey(key: string): string {
  const trimmed = key.trim();
  if (!trimmed) return '';
  if (trimmed.length <= 8) return '****';
  return `${trimmed.slice(0, 4)}••••${trimmed.slice(-4)}`;
}

export function renderPromptTemplate(
  template: string,
  vars: { selection?: string; context?: string }
): string {
  return template
    .replace(/\{\{selection\}\}/g, vars.selection ?? '')
    .replace(/\{\{context\}\}/g, vars.context ?? '');
}
