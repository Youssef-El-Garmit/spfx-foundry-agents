/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../FoundryAgentsWebPartStrings.d.ts" />
/**
 * Types for server-side MCP tools run by Azure AI Foundry (OAuth identity passthrough) and for Word
 * documents produced by the doc_* tools of the Microsoft 365 MCP.
 */
import strings from 'FoundryAgentsWebPartStrings';

/** Connection state of one MCP server for the signed-in user. */
export type McpConnectionStatus = 'connected' | 'consent_required' | 'error' | 'unknown';

/** MCP server declared on the agent (definition.tools[type=mcp]). */
export interface IMcpServerInfo {
  /** server_label in the agent definition (e.g. "ms365", "crm"). */
  label: string;
  url?: string;
  connection?: string;
}

/** Live state of an MCP server for the current user. */
export interface IMcpServerState extends IMcpServerInfo {
  status: McpConnectionStatus;
  /** Number of tools listed by the server on the last call (0 when not connected). */
  toolCount?: number;
  /** Consent link returned by Foundry when the user must sign in. */
  consentLink?: string;
  error?: string;
  /** Epoch ms of the last observation. */
  checkedAt?: number;
}

/** Sign-in request raised by Foundry during a turn (item type oauth_consent_request). */
export interface IConsentRequest {
  serverLabel: string;
  link: string;
  /** Set once the user completed the sign-in popup. */
  done?: boolean;
}

/** Word document created, read or edited by a doc_* tool. */
export interface IDocumentRef {
  driveId: string;
  itemId: string;
  name: string;
  webUrl?: string;
  /** Epoch ms of the last change made by the agent (used to refresh the preview). */
  updatedAt?: number;
}

/** MCP tool call observed in the stream. */
export interface IMcpCallEvent {
  serverLabel: string;
  name: string;
  arguments?: string;
  output?: string;
  error?: string;
}

/** Extract the document reference from a doc_* tool output (JSON with a `document` field). */
export function documentFromToolOutput(name: string, output: string | undefined): IDocumentRef | undefined {
  if (!output || !/^doc_(create|edit|read)$/.test(name)) return undefined;
  try {
    const data = JSON.parse(output) as { document?: { driveId?: string; itemId?: string; name?: string; webUrl?: string } };
    const d = data?.document;
    if (!d?.driveId || !d.itemId || !d.name) return undefined;
    // Templates are read too: only show documents the agent is working on (not files from a "...Templates" library).
    if (name === 'doc_read' && /\/[^/?#]*templates?\//i.test(d.webUrl ?? '')) return undefined;
    return { driveId: d.driveId, itemId: d.itemId, name: d.name, webUrl: d.webUrl, updatedAt: name === 'doc_read' ? undefined : Date.now() };
  } catch {
    return undefined;
  }
}

/** Readable fallback for an unknown tool name (e.g. "list_mail_messages" → "List mail messages"). */
function prettifyToolName(name: string): string {
  const words = name.replace(/([a-z0-9])([A-Z])/g, '$1 $2').replace(/[-_.]+/g, ' ').trim().toLowerCase();
  return words ? words.charAt(0).toUpperCase() + words.slice(1) : name;
}

/** Friendly localized label for an MCP tool (loader text). */
export function friendlyToolLabel(name: string): string {
  const map: Record<string, string> = {
    search_tools: strings.ToolSearchTools,
    get_tool_schema: strings.ToolGetToolSchema,
    call_tool: strings.ToolCallTool,
    doc_list_templates: strings.ToolDocListTemplates,
    doc_list_content_blocks: strings.ToolDocListContentBlocks,
    doc_create: strings.ToolDocCreate,
    doc_read: strings.ToolDocRead,
    doc_edit: strings.ToolDocEdit
  };
  return map[name] ?? prettifyToolName(name);
}
