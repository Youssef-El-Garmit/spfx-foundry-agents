/**
 * Azure AI Foundry Responses API (openai/responses).
 * Helpers for response.output / response.data items (function_call detection, text extraction)
 * and createResponse (POST) for the generic tool loop in AgentStreamManager.
 */

import type { GetTokenFn } from './SpFxTokenCredential';
import type { IResponseOutputItemLike } from './foundryApiTypes';

const API_VERSION = '2025-11-15-preview';

export interface IResponsesApiConfig {
  endpoint: string;
  getToken: GetTokenFn;
  agentName: string;
  /** Tools to send in the request (optional; some backends use agent's default tools). */
  tools?: Array<{ type: 'function'; function: { name: string; description?: string; parameters?: unknown } }>;
}

export interface IResponsesApiInputItem {
  type?: string;
  role?: string;
  content?: string | Array<{ type?: string; text?: string }>;
}

/** Response from POST responses (non-stream). Items may be in .output or .data (list API). */
export interface IResponsesApiResponse {
  id: string;
  status?: string;
  output?: IResponseOutputItemLike[];
  data?: IResponseOutputItemLike[];
  usage?: { total_tokens?: number };
}

function normalizeEndpoint(url: string): string {
  return url.replace(/\/$/, '');
}

/** Detect function_call items from items array (response.data or response.output). */
export function detectToolCallsFromResponse(response: IResponsesApiResponse): Array<{ callId: string; name: string; arguments: string }> {
  const items = response?.data ?? response?.output ?? [];
  const result: Array<{ callId: string; name: string; arguments: string }> = [];
  for (const item of items) {
    const t = (item.type ?? item.type_) ?? '';
    if (t !== 'function_call') continue;
    const callId = item.call_id ?? '';
    const name = (item.name ?? item.function_name ?? '').trim();
    const args = (item.arguments ?? item.function_arguments ?? '').trim();
    if (callId && name) {
      result.push({ callId, name, arguments: args });
    }
  }
  return result;
}

/** Extract final assistant text from items array (response.data or response.output). */
export function getTextFromResponseOutput(response: IResponsesApiResponse): string {
  const items = response?.data ?? response?.output ?? [];
  const parts: string[] = [];
  for (const item of items) {
    const t = (item.type ?? item.type_) ?? '';
    if (t === 'message' && Array.isArray((item as { content?: Array<{ type?: string; text?: string }> }).content)) {
      for (const c of (item as { content: Array<{ type?: string; text?: string }> }).content) {
        if (c?.type === 'output_text' && c.text) parts.push(c.text);
      }
    }
    if (t === 'output_text' && (item as { text?: string }).text) {
      parts.push((item as { text: string }).text);
    }
  }
  return parts.join('\n').trim();
}

/**
 * Create a response (POST). Returns full response body.
 * Exported for use by FoundryAgentService when continuing after stream with tool outputs.
 */
export async function createResponse(
  endpoint: string,
  token: string,
  body: {
    agent?: { name: string; type: string };
    model?: string;
    input: string | IResponsesApiInputItem[];
    previous_response_id?: string;
    tools?: unknown[];
  }
): Promise<IResponsesApiResponse> {
  const base = normalizeEndpoint(endpoint);
  const url = `${base}/openai/responses?api-version=${API_VERSION}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Responses API error ${res.status}: ${errText}`);
  }
  return (await res.json()) as IResponsesApiResponse;
}

