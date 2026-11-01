/**
 * Strip injected [Context: ...] prefix from message or title text for display.
 * Used so user messages and conversation titles don't show the context block.
 */
export function stripContextPrefix(content: string): string {
  const match = content.match(/^\[Context:[\s\S]*?\]\s*\n+\s*/);
  return match ? content.slice(match[0].length).trim() : content;
}
