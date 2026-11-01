/**
 * SSE stream handler + function_call collector + submit loop.
 * Uses FoundryAgentService for all API access (conversations, stream start, response detail, tool continuation).
 * Ensures every function_call is completed with tool output before the turn ends, so the next user message
 * does not get 400 "No tool output found for function call".
 */

import {
  getTextFromResponseOutput,
  type IResponsesApiResponse
} from './FoundryResponsesApi';
import {
  executeFromFunctionCall,
  type GetTokenFn as FunctionExecutorGetTokenFn
} from './FunctionExecutor';
import type { FoundryAgentService } from './FoundryAgentService';
import type { IConsentRequest, IMcpCallEvent } from './mcpTypes';
import type { AgentStep } from '../types/chat';

export type AgentStreamState = 'idle' | 'thinking' | 'executing' | 'done';

export interface IAgentStreamCallbacks {
  onStateChange: (state: AgentStreamState, payload?: { functionNames?: string[]; text?: string }) => void;
  onChunk: (text: string) => void;
  /** Conversation id known (created before the first message is sent). */
  onConversation?: (conversationId: string) => void;
  /** Foundry asks the user to sign in to an MCP server (OAuth identity passthrough). */
  onConsentRequest?: (consent: IConsentRequest) => void;
  /** Tools listed by an MCP server for this user (0 tools + consent = not connected). */
  onMcpListTools?: (serverLabel: string, toolCount: number, error?: string) => void;
  /** Server-side MCP tool call started (added) or finished (done, with output). */
  onMcpCall?: (call: IMcpCallEvent, phase: 'start' | 'done') => void;
  /** A step (reasoning summary or tool action) was created or updated (same id = replace). */
  onStep?: (step: AgentStep) => void;
}

const DEFAULT_BLOCKED_HTTP_METHODS = ['DELETE', 'PATCH'];

function getBlockedMethods(blockedHttpMethods: string[] | undefined): string[] {
  if (!blockedHttpMethods || blockedHttpMethods.length === 0) return [...DEFAULT_BLOCKED_HTTP_METHODS];
  return blockedHttpMethods.filter((m): m is string => typeof m === 'string').map((m) => m.toUpperCase());
}

function parseMethodFromArguments(argumentsJson: string): string | undefined {
  try {
    const data = argumentsJson && argumentsJson.trim() ? JSON.parse(argumentsJson) : {};
    const method = data?.request?.method;
    return typeof method === 'string' ? method.toUpperCase() : undefined;
  } catch {
    return undefined;
  }
}

export interface IAgentStreamManagerConfig {
  /** Service for conversations, stream start, getResponseDetail, submitToolOutputs. */
  service: FoundryAgentService;
  /** Resolved agent name for API (e.g. from agents.get(agentId).name) */
  agentNameForApi: string;
  tenantHost: string;
  callbacks: IAgentStreamCallbacks;
  signal?: AbortSignal;
  /** When false (default), function_call are not executed; model receives an error. */
  allowFunctionCalls?: boolean;
  /** HTTP methods that must not be executed. Default when undefined: ['DELETE', 'PATCH']. */
  blockedHttpMethods?: string[];
  /** Optional site context for conversation metadata (creation_site_url). */
  siteContext?: { siteUrl: string };
  /** Signed-in user (UPN): stored as the conversation owner so the history lists only the user's own conversations. */
  ownerUpn?: string;
  /** Optional context string to prefix to the user message (e.g. [Context: ...]). */
  additionalContextPrefix?: string;
  /** Builds the context prefix once the conversation id is known (takes precedence over additionalContextPrefix). */
  buildContextPrefix?: (conversationId: string) => string | undefined;
}

/** One function_call item from response.output. */
export interface ICollectedFunctionCall {
  call_id: string;
  name: string;
  arguments: string;
}

/**
 * Extract a display-friendly message from stream error payloads.
 * API may send error.message as a JSON string (e.g. tool_user_error); try to parse and use inner message.
 */
function getStreamErrorMessage(payload: { error?: { message?: string; code?: string }; response?: { error?: { message?: string; code?: string } } }): string {
  const msg =
    typeof payload.error?.message === 'string'
      ? payload.error.message
      : typeof payload.response?.error?.message === 'string'
        ? payload.response.error.message
        : '';
  if (!msg) return 'Stream error';
  try {
    const parsed = JSON.parse(msg) as { message?: string; error?: string };
    if (typeof parsed.message === 'string' && parsed.message.length > 0) return parsed.message;
    if (typeof parsed.error === 'string' && parsed.error.length > 0) return parsed.error;
  } catch {
    // not JSON, use raw
  }
  return msg.length > 500 ? msg.slice(0, 500) + '…' : msg;
}

/**
 * Parse SSE stream: yield parsed events { type, data }.
 */
async function* parseSSE(reader: ReadableStreamDefaultReader<Uint8Array>): AsyncGenerator<{ type: string; data: string }> {
  const decoder = new TextDecoder();
  let buffer = '';
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split(/\r?\n/);
    buffer = lines.pop() ?? '';
    let eventType = '';
    for (const line of lines) {
      if (line.startsWith('event:')) {
        eventType = line.slice(6).trim();
      } else if (line.startsWith('data:') && eventType) {
        const data = line.slice(5).trim();
        if (data === '[DONE]') return;
        yield { type: eventType, data };
        eventType = '';
      }
    }
  }
}

/**
 * Extract function_call items from the items array (response.data or response.output).
 */
function collectFunctionCallsFromResponse(response: {
  data?: Array<{ type?: string; call_id?: string; name?: string; arguments?: string }>;
  output?: Array<{ type?: string; call_id?: string; name?: string; arguments?: string }>;
}): ICollectedFunctionCall[] {
  const items = response?.data ?? response?.output ?? [];
  const result: ICollectedFunctionCall[] = [];
  for (const item of items) {
    const t = item.type ?? '';
    if (t !== 'function_call') continue;
    const call_id = item.call_id ?? '';
    const name = (item.name ?? '').trim();
    const args = (item.arguments ?? '').trim();
    if (call_id && name) {
      result.push({ call_id, name, arguments: args });
    }
  }
  return result;
}

export class AgentStreamManager {
  private readonly _config: IAgentStreamManagerConfig;

  constructor(config: IAgentStreamManagerConfig) {
    this._config = config;
  }

  /** One request per turn: message as input + conversation. */
  private async *openStreamWithInput(conversationId: string, message: string): AsyncGenerator<{ eventType: string; data: unknown; textDelta?: string }> {
    const { service, agentNameForApi, signal } = this._config;
    const res = await service.startResponseStreamWithInput(conversationId, message, agentNameForApi, signal);
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Start stream failed: ${res.status} ${err}`);
    }
    const reader = res.body?.getReader();
    if (!reader) throw new Error('No response body');
    yield* this.readSSE(reader);
  }

  /**
   * Open SSE stream by continuing from a previous response (next user turn).
   * Use when previous turn is complete so the server does not see a pending tool call.
   */
  private async *openStreamWithContinuation(
    previousResponseId: string,
    userMessage: string
  ): AsyncGenerator<{ eventType: string; data: unknown; textDelta?: string }> {
    const { service, agentNameForApi, signal } = this._config;
    const res = await service.startResponseStreamWithContinuation(
      previousResponseId,
      userMessage,
      agentNameForApi,
      signal
    );
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Start stream failed: ${res.status} ${err}`);
    }
    const reader = res.body?.getReader();
    if (!reader) throw new Error('No response body');
    yield* this.readSSE(reader);
  }

  private async *readSSE(reader: ReadableStreamDefaultReader<Uint8Array>): AsyncGenerator<{ eventType: string; data: unknown; textDelta?: string }> {
    for await (const { type: eventType, data: dataStr } of parseSSE(reader)) {
      if (this._config.signal?.aborted) throw new DOMException('Aborted', 'AbortError');
      let data: unknown;
      try {
        data = dataStr ? (JSON.parse(dataStr) as unknown) : {};
      } catch {
        data = {};
      }
      if (eventType === 'error' || eventType === 'response.failed') {
        const message = getStreamErrorMessage(data as Parameters<typeof getStreamErrorMessage>[0]);
        throw new Error(message);
      }
      const obj = data as { type?: string; delta?: string; response?: IResponsesApiResponse };
      if (eventType === 'response.output_item.added' || eventType === 'response.output_item.done') {
        this.handleOutputItem((data as { item?: Record<string, unknown> }).item, eventType === 'response.output_item.done');
      } else if (eventType === 'response.reasoning_summary_text.delta' || eventType === 'response.reasoning_summary_part.added') {
        const d = data as { item_id?: string; delta?: string };
        const step = d.item_id ? this._steps.get(d.item_id) : undefined;
        if (step && step.kind === 'thinking') {
          const add = eventType === 'response.reasoning_summary_part.added' ? (step.text ? '\n\n' : '') : (d.delta ?? '');
          if (add) this.emitStep({ ...step, text: step.text + add });
        }
      }
      if (eventType === 'response.output_text.delta' && typeof obj.delta === 'string') {
        yield { eventType, data, textDelta: obj.delta };
      } else {
        yield { eventType, data };
      }
    }
  }

  private _consent: IConsentRequest | undefined;
  private _steps = new Map<string, AgentStep>();
  private _sawFunctionCall = false;

  private emitStep(step: AgentStep): void {
    this._steps.set(step.id, step);
    this._config.callbacks.onStep?.(step);
  }

  /** Text of a reasoning item's summary (when Foundry sends it only on the done event). */
  private static summaryText(item: Record<string, unknown>): string {
    const s = item.summary;
    if (typeof s === 'string') return s;
    if (Array.isArray(s)) return s.map((x: { text?: string }) => x?.text ?? '').filter(Boolean).join('\n\n');
    return '';
  }

  /** Server-side items (MCP tools, sign-in requests) seen in the stream. */
  private handleOutputItem(item: Record<string, unknown> | undefined, done: boolean): void {
    if (!item) return;
    const cb = this._config.callbacks;
    const type = String(item.type ?? '');
    const label = String(item.server_label ?? '');
    const id = String((type === 'function_call' ? item.call_id : item.id) ?? item.id ?? '');
    const now = Date.now();
    if (type === 'reasoning' && id) {
      const prev = this._steps.get(id);
      const text = prev?.kind === 'thinking' && prev.text ? prev.text : AgentStreamManager.summaryText(item);
      this.emitStep({ kind: 'thinking', id, text, startedAt: prev?.startedAt ?? now, endedAt: done ? now : undefined });
    } else if ((type === 'mcp_call' || type === 'function_call') && id) {
      if (type === 'function_call') this._sawFunctionCall = true;
      const prev = this._steps.get(id);
      const err = item.error ? (typeof item.error === 'string' ? item.error : JSON.stringify(item.error)) : undefined;
      const output = typeof item.output === 'string' ? item.output : undefined;
      this.emitStep({
        kind: 'tool',
        id,
        server: type === 'mcp_call' ? label : 'client',
        name: String(item.name ?? ''),
        status: !done || type === 'function_call' ? 'running' : err || /^ERREUR/.test(output ?? '') ? 'error' : 'done',
        arguments: typeof item.arguments === 'string' ? item.arguments : undefined,
        output,
        error: err,
        startedAt: prev?.startedAt ?? now,
        endedAt: done && type === 'mcp_call' ? now : undefined
      });
    }
    if (type === 'oauth_consent_request' && typeof item.consent_link === 'string') {
      if (!this._consent) {
        this._consent = { serverLabel: label, link: item.consent_link };
        cb.onConsentRequest?.(this._consent);
      }
    } else if (type === 'mcp_list_tools' && done) {
      const tools = Array.isArray(item.tools) ? item.tools.length : 0;
      cb.onMcpListTools?.(label, tools, item.error ? String(item.error) : undefined);
    } else if (type === 'mcp_call') {
      const name = String(item.name ?? '');
      if (!done) cb.onStateChange('executing', { functionNames: [name] });
      cb.onMcpCall?.({
        serverLabel: label,
        name,
        arguments: typeof item.arguments === 'string' ? item.arguments : undefined,
        output: typeof item.output === 'string' ? item.output : undefined,
        error: item.error ? (typeof item.error === 'string' ? item.error : JSON.stringify(item.error)) : undefined
      }, done ? 'done' : 'start');
      if (done) cb.onStateChange('thinking');
    }
  }

  /**
   * Run the full flow. For the next user message in an existing conversation, pass lastResponseId
   * so the stream is started with previous_response_id (avoids 400 "No tool output found for function call").
   * Returns lastResponseId so the client can pass it on the next message.
   */
  async run(
    conversationId: string | undefined,
    userMessage: string,
    lastResponseId?: string,
    seedItems?: Array<{ role: 'user' | 'assistant'; text: string }>
  ): Promise<{ conversationId: string; outputText: string; lastResponseId?: string; consent?: IConsentRequest }> {
    const { service, agentNameForApi, tenantHost, callbacks, additionalContextPrefix, buildContextPrefix } = this._config;
    const getToken: FunctionExecutorGetTokenFn = service.getTokenFn();
    this._consent = undefined;
    const withPrefix = (convId: string): string => {
      const prefix = (buildContextPrefix ? buildContextPrefix(convId) : additionalContextPrefix)?.trim();
      return prefix ? prefix + '\n\n' + userMessage : userMessage;
    };

    let convId: string;
    let streamSource: AsyncGenerator<{ eventType: string; data: unknown; textDelta?: string }>;

    this._steps.clear();
    this._sawFunctionCall = false;
    if (!conversationId) {
      // Conversation created first so its id can be given to the agent (documents remember it for the Word add-in).
      convId = await service.createConversation(this._config.siteContext, userMessage, {
        agent: agentNameForApi,
        owner: this._config.ownerUpn,
        items: seedItems
      });
      callbacks.onConversation?.(convId);
      streamSource = this.openStreamWithInput(convId, withPrefix(convId));
    } else if (lastResponseId && this._config.allowFunctionCalls === true) {
      convId = conversationId;
      streamSource = this.openStreamWithContinuation(lastResponseId, withPrefix(conversationId));
    } else {
      convId = conversationId;
      streamSource = this.openStreamWithInput(convId, withPrefix(convId));
    }

    callbacks.onStateChange('thinking');

    let outputText = '';
    let completedResponseId: string | undefined;

    // Phase 1: consume SSE until response.completed
    for await (const chunk of streamSource) {
      if (chunk.textDelta) {
        outputText += chunk.textDelta;
        callbacks.onChunk(chunk.textDelta);
      }
      const data = chunk.data as { type?: string; response?: { id?: string } };
      if (chunk.eventType === 'response.completed' || data.type === 'response.completed') {
        const resp = data.response;
        if (resp?.id) completedResponseId = resp.id;
      }
    }

    // Server-side tools (MCP) are fully handled by Foundry: no extra request unless a client function was called.
    if (!completedResponseId || !this._sawFunctionCall) {
      callbacks.onStateChange('done', { text: outputText });
      return { conversationId: convId, outputText, lastResponseId: completedResponseId, consent: this._consent };
    }

    // Phase 2: fetch full response so we never miss function_call items
    let fullResponse: IResponsesApiResponse = await service.getResponseDetail(completedResponseId) as IResponsesApiResponse;

    // Phase 3: tool loop — submit tool outputs until no more function_calls
    while (fullResponse) {
      const functionCalls = collectFunctionCallsFromResponse(fullResponse);
      if (functionCalls.length === 0) {
        const text = getTextFromResponseOutput(fullResponse);
        if (text) {
          outputText = text;
          callbacks.onChunk(text);
        }
        callbacks.onStateChange('done', { text: outputText });
        return { conversationId: convId, outputText, lastResponseId: fullResponse.id, consent: this._consent };
      }

      callbacks.onStateChange('executing', { functionNames: functionCalls.map(f => f.name) });

      const allowFunctionCalls = this._config.allowFunctionCalls === true;
      const blockedMethods = getBlockedMethods(this._config.blockedHttpMethods);

      const functionOutputs: Array<{ callId: string; output: string }> = [];
      for (const fc of functionCalls) {
        if (!allowFunctionCalls) {
          functionOutputs.push({
            callId: fc.call_id,
            output: JSON.stringify({
              error: 'Function calls are disabled for this agent in this host (client-side tools such as Office.js only run in a host that implements them). Use the agent\'s server-side tools instead.'
            })
          });
          continue;
        }
        const method = parseMethodFromArguments(fc.arguments);
        if (method && blockedMethods.includes(method)) {
          functionOutputs.push({
            callId: fc.call_id,
            output: JSON.stringify({ error: `HTTP method ${method} is blocked for this agent.` })
          });
          continue;
        }
        try {
          const result = await executeFromFunctionCall(fc.name, fc.arguments, getToken, tenantHost);
          const output = typeof result === 'string' ? result : JSON.stringify(result);
          functionOutputs.push({ callId: fc.call_id, output });
        } catch (e) {
          const errMsg = e instanceof Error ? e.message : String(e);
          functionOutputs.push({ callId: fc.call_id, output: JSON.stringify({ error: errMsg }) });
        }
      }

      for (const fo of functionOutputs) {
        const st = this._steps.get(fo.callId);
        if (st && st.kind === 'tool') this.emitStep({ ...st, status: /"error"/.test(fo.output) ? 'error' : 'done', output: fo.output, endedAt: Date.now() });
      }
      fullResponse = await service.submitToolOutputsAndGetNextResponse(
        agentNameForApi,
        fullResponse.id,
        functionOutputs
      );
    }

    callbacks.onStateChange('done', { text: outputText });
    return { conversationId: convId, outputText };
  }
}
