/**
 * Azure AI Foundry Agent service built on @azure/ai-projects.
 * Uses AIProjectClient → getOpenAIClient() → conversations + responses APIs.
 * Files attached by the user are uploaded via openAIClient.files before the
 * message is sent; the uploaded file IDs are included in the message text so
 * the agent can reference them (e.g. via Code Interpreter or File Search tools).
 */
import { AIProjectClient } from '@azure/ai-projects';
import { AZURE_AI_SCOPE, type GetTokenFn } from './SpFxTokenCredential';
import { SpFxTokenCredential } from './SpFxTokenCredential';
import { createResponse, type IResponsesApiResponse } from './FoundryResponsesApi';
import type { AgentStep, IChatMessage } from '../types/chat';
import { documentFromToolOutput, type IDocumentRef, type IMcpServerInfo, type IMcpServerState } from './mcpTypes';

/** Context block prefixed to user messages ("[Contexte] …" or legacy "[Context: …]"), hidden when displaying history. */
export function stripContextPrefix(text: string): string {
  let out = text.replace(/^\[(Contexte|Context)[^\n]*\]?\s*\n\n/, '').replace(/^\[Context:[\s\S]*?\]\s*\n\n/, '');
  // Blocks added by the Word add-in (<user_context>, <doc_context>, <attachments>) when the conversation continues in Word.
  let prev: string;
  do {
    prev = out;
    out = out.replace(/^<(\w+_context|attachments)>\n[\s\S]*?\n<\/\1>\n\n/, '');
  } while (out !== prev);
  return out;
}

const CONVERSATIONS_API_VERSION = '2025-11-15-preview';
const TITLE_MAX_LEN = 80;

function normalizeEndpoint(url: string): string {
  return url.replace(/\/$/, '');
}

function cropTitle(text: string): string {
  const s = text.replace(/\s+/g, ' ').trim();
  return s.length <= TITLE_MAX_LEN ? s : s.slice(0, TITLE_MAX_LEN - 3) + '...';
}

/** SDK metadata may expect string values; normalize for update. */
function metadataForUpdate(meta: Record<string, unknown>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(meta)) {
    out[k] = (v === undefined || v === null) ? '' : typeof v === 'string' ? v : String(v);
  }
  return out;
}

/** The conversation no longer exists in Foundry (404 / conversation_not_found). */
export class ConversationGoneError extends Error {
  constructor(public readonly conversationId: string) {
    super(`Conversation ${conversationId} introuvable (conversation_not_found)`);
    this.name = 'ConversationGoneError';
    // SPFx compiles to ES5: restore the prototype, otherwise instanceof fails on Error subclasses.
    Object.setPrototypeOf(this, ConversationGoneError.prototype);
  }
}

/** True for Foundry's "conversation not found" errors (HTTP 404, or 400 conversation_not_found when answering). */
export function isConversationGone(err: unknown): boolean {
  if (err instanceof ConversationGoneError || (err as Error)?.name === 'ConversationGoneError') return true;
  const msg = err instanceof Error ? err.message : String(err ?? '');
  return /conversation_not_found|Conversation with id '[^']+' not found/i.test(msg);
}

// ── Public types ──────────────────────────────────────────────────────────────

export interface IConversationSummary {
  id: string;
  title: string;
  updatedAt?: Date;
  /** From metadata.creation_site_url */
  creationSiteUrl?: string;
  /** From metadata.last_used_site_url */
  lastUsedSiteUrl?: string;
  /** From metadata.last_item_id – for notification badge */
  lastItemId?: string;
  /** From metadata.last_seen_item_id */
  lastSeenItemId?: string;
  /** Full metadata as listed (kept when metadata is rewritten, so agent/owner are not lost). */
  rawMeta?: Record<string, unknown>;
}

export interface ListConversationsOptions {
  limit?: number;
  after?: string;
  /** Only return conversations that have a title in metadata (started from this web part). */
  onlyWithTitleInMeta?: boolean;
  /** Signed-in user's UPN: conversations owned by someone else (metadata.owner) are left out. */
  ownerUpn?: string;
}

export interface ListConversationsResult {
  items: IConversationSummary[];
  nextCursor?: string;
}

/** A single annotation on a conversation item content part */
export interface IItemAnnotation {
  type: string;
  text?: string;
  url?: string;
  title?: string;
  start_index?: number;
  end_index?: number;
  file_id?: string;
  filename?: string;
  index?: number;
  quote?: string;
}

/** A content part inside a conversation item */
export interface IItemContentPart {
  type: string;
  text?: string;
  refusal?: string;
  /** Summary of a reasoning step (string or array of summary blocks) */
  summary?: string | Array<{ type: string; text: string }>;
  annotations?: IItemAnnotation[];
}

/** Full detail of a single conversation item */
export interface IConversationItemDetail {
  id: string;
  type: string;
  role?: string;
  status?: string;
  created_at?: number;
  content?: IItemContentPart[];
  created_by?: {
    type?: string;
    response_id?: string;
  };
}

/** A single item in a response's output array */
export interface IResponseOutputItem {
  type: string;
  id?: string;
  role?: string;
  status?: string;
  content?: Array<{
    type: string;
    text?: string;
    annotations?: IItemAnnotation[];
  }>;
  /** For reasoning items */
  summary?: string | Array<{ type: string; text: string }>;
  /** For file_search_call / web_search_call */
  queries?: string[];
  results?: Array<{
    file_id?: string;
    filename?: string;
    score?: number;
    text?: string;
    url?: string;
    title?: string;
  }>;
  /** For function_call */
  call_id?: string;
  name?: string;
  arguments?: string;
}

/** Full response detail from GET /openai/responses/{responseId} */
export interface IResponseDetail {
  id: string;
  model?: string;
  status?: string;
  created_at?: number;
  output?: IResponseOutputItem[];
  usage?: {
    input_tokens?: number;
    output_tokens?: number;
    total_tokens?: number;
    input_tokens_details?: {
      cached_tokens?: number;
      text_tokens?: number;
    };
    output_tokens_details?: {
      reasoning_tokens?: number;
    };
  };
}

/** Combined message details returned by getMessageDetails() */
export interface IMessageDetailsData {
  item: IConversationItemDetail;
  response?: IResponseDetail;
}

// ── Internal types ────────────────────────────────────────────────────────────

type OpenAIClient = Awaited<ReturnType<AIProjectClient['getOpenAIClient']>>;

interface FoundryConversationItem {
  id: string;
  created_at?: number;
  metadata?: Record<string, unknown>;
}

// ── Service ───────────────────────────────────────────────────────────────────

export class FoundryAgentService {
  private readonly _client: AIProjectClient;
  private readonly _endpoint: string;
  private readonly _getTokenFn: GetTokenFn;
  private _openAIClient: Promise<OpenAIClient> | undefined;

  constructor(projectEndpoint: string, getTokenFn: GetTokenFn) {
    this._endpoint = normalizeEndpoint(projectEndpoint);
    this._client = new AIProjectClient(this._endpoint, new SpFxTokenCredential(getTokenFn));
    this._getTokenFn = getTokenFn;
  }

  // ── Private helpers ─────────────────────────────────────────────────────────

  private openAI(): Promise<OpenAIClient> {
    if (!this._openAIClient) this._openAIClient = this._client.getOpenAIClient();
    return this._openAIClient;
  }

  // ── Public API ──────────────────────────────────────────────────────────────

  /**
   * Upload a single file to Azure AI Projects storage.
   * Returns the file_id assigned by the service.
   * Called immediately when the user adds an attachment — before sending the message.
   */
  async uploadFile(file: File): Promise<string> {
    const openAI = await this.openAI();
    const uploaded = await openAI.files.create({ file, purpose: 'assistants' });
    return uploaded.id;
  }

  /**
   * Resolve agent ID to the name required by the Responses API (agent.name).
   */
  async getResolvedAgentName(agentId: string): Promise<string> {
    const agent = await this._client.agents.get(agentId);
    return agent.name;
  }

  /** Get a token for the Azure AI scope. Used by stream/tool flow. */
  async getToken(): Promise<string> {
    return this._getTokenFn(AZURE_AI_SCOPE);
  }

  /** Token getter by resource (for FunctionExecutor). */
  getTokenFn(): GetTokenFn {
    return this._getTokenFn;
  }

  /**
   * Create a conversation and add the first user message (Responses API flow).
   * Sets metadata.title (and optionally creation_site_url) so the conversation appears in history when showOnlyWebPartConversations is true.
   * @param userMessage - Full message content sent to the API (may include injected context prefix).
   * @param siteContext - Optional site URL for metadata.creation_site_url.
   * @param titleForMetadata - If provided, used for metadata.title only (e.g. raw user message without context). Use this so the conversation title never contains the context block.
   */
  async createConversationWithMessage(
    userMessage: string,
    siteContext?: { siteUrl: string },
    titleForMetadata?: string
  ): Promise<string> {
    const openAI = await this.openAI();
    const conv = await openAI.conversations.create({
      items: [{ type: 'message', role: 'user', content: userMessage }]
    });
    const title = cropTitle(titleForMetadata !== undefined && titleForMetadata !== '' ? titleForMetadata : userMessage);
    const metadata: Record<string, unknown> = {};
    if (title) metadata.title = title;
    const siteUrl = siteContext?.siteUrl?.trim();
    if (siteUrl) metadata.creation_site_url = siteUrl;
    if (Object.keys(metadata).length > 0) {
      try {
        await openAI.conversations.update(conv.id, { metadata: metadataForUpdate(metadata) });
      } catch {
        // Non-fatal: conversation was created, history filter may not show it until next load
      }
    }
    return conv.id;
  }

  /**
   * Create an empty conversation (metadata title / creation site) and return its id.
   * The first message is added afterwards so that it can carry the conversation id in its context block.
   */
  async createConversation(
    siteContext?: { siteUrl: string },
    titleForMetadata?: string,
    extra?: { agent?: string; owner?: string; items?: Array<{ role: 'user' | 'assistant'; text: string }> }
  ): Promise<string> {
    const openAI = await this.openAI();
    const metadata: Record<string, unknown> = {};
    // agent + owner: the history lists the conversation even if its first answer failed (Foundry only links a
    // conversation to an agent once an answer completed), and only for its owner.
    if (extra?.agent) metadata.agent = extra.agent.replace(/-word$/, '');
    if (extra?.owner) metadata.owner = extra.owner.toLowerCase();
    const title = cropTitle(titleForMetadata ?? '');
    if (title) metadata.title = title;
    const siteUrl = siteContext?.siteUrl?.trim();
    if (siteUrl) metadata.creation_site_url = siteUrl;
    // Optional earlier exchanges (a conversation recreated after Foundry lost the original one).
    const items = (extra?.items ?? []).filter((m) => m.text.trim()).map((m) => ({
      type: 'message' as const,
      role: m.role,
      content: [{ type: m.role === 'user' ? 'input_text' as const : 'output_text' as const, text: m.text }]
    }));
    const conv = await openAI.conversations.create({
      ...(Object.keys(metadata).length ? { metadata: metadataForUpdate(metadata) } : {}),
      ...(items.length ? { items } : {})
    } as Parameters<typeof openAI.conversations.create>[0]);
    return conv.id;
  }

  /** MCP servers declared on the agent's latest version (definition.tools of type "mcp"). */
  async getAgentMcpServers(agentName: string): Promise<IMcpServerInfo[]> {
    const token = await this.getToken();
    const res = await fetch(`${this._endpoint}/agents/${encodeURIComponent(agentName)}?api-version=${CONVERSATIONS_API_VERSION}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) return [];
    const body = (await res.json()) as { versions?: { latest?: { definition?: { tools?: Array<Record<string, unknown>> } } } };
    const tools = body.versions?.latest?.definition?.tools ?? [];
    return tools
      .filter((t) => t.type === 'mcp' && typeof t.server_label === 'string')
      .map((t) => ({ label: String(t.server_label), url: t.server_url as string | undefined, connection: t.project_connection_id as string | undefined }));
  }

  /**
   * Check the MCP connections of the current user with a minimal, unstored response (tool_choice none):
   * Foundry lists each server's tools for this user and returns a sign-in link for a server not yet connected.
   */
  async probeMcp(agentName: string, servers: IMcpServerInfo[]): Promise<IMcpServerState[]> {
    const token = await this.getToken();
    const res = await fetch(`${this._endpoint}/openai/responses?api-version=${CONVERSATIONS_API_VERSION}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        agent: { name: agentName, type: 'agent_reference' },
        input: 'ping',
        tool_choice: 'none',
        store: false,
        max_output_tokens: 16
      })
    });
    const now = Date.now();
    if (!res.ok) {
      const err = await res.text();
      return servers.map((s) => ({ ...s, status: 'error' as const, error: err.slice(0, 300), checkedAt: now }));
    }
    const body = (await res.json()) as { output?: Array<Record<string, unknown>> };
    const out = body.output ?? [];
    return servers.map((s) => {
      const list = out.find((i) => i.type === 'mcp_list_tools' && i.server_label === s.label);
      const consent = out.find((i) => i.type === 'oauth_consent_request' && i.server_label === s.label);
      const toolCount = Array.isArray(list?.tools) ? (list?.tools as unknown[]).length : 0;
      if (consent) return { ...s, status: 'consent_required' as const, consentLink: String(consent.consent_link), toolCount: 0, checkedAt: now };
      if (list?.error) return { ...s, status: 'error' as const, error: String(list.error), toolCount, checkedAt: now };
      if (toolCount > 0) return { ...s, status: 'connected' as const, toolCount, checkedAt: now };
      // Listed with no tools and no link: Foundry only returns one sign-in link per response.
      return { ...s, status: list ? 'consent_required' as const : 'unknown' as const, toolCount, checkedAt: now };
    });
  }

  /** Word documents created or edited in a conversation (doc_* tool outputs), most recent last. */
  async getConversationDocuments(conversationId: string): Promise<IDocumentRef[]> {
    try {
      const token = await this.getToken();
      const url = new URL(`${this._endpoint}/openai/conversations/${conversationId}/items`);
      url.searchParams.set('api-version', CONVERSATIONS_API_VERSION);
      url.searchParams.set('limit', '100');
      url.searchParams.set('order', 'asc');
      const res = await fetch(url.toString(), { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) return [];
      const body = (await res.json()) as { data?: Array<Record<string, unknown>> };
      const docs = new Map<string, IDocumentRef>();
      for (const item of body.data ?? []) {
        if (item.type !== 'mcp_call') continue;
        const doc = documentFromToolOutput(String(item.name ?? ''), typeof item.output === 'string' ? item.output : undefined);
        if (doc) { docs.delete(doc.itemId); docs.set(doc.itemId, { ...doc, updatedAt: undefined }); }
      }
      return Array.from(docs.values());
    } catch {
      return [];
    }
  }

  /**
   * Add a user message to an existing conversation.
   * Only call after the previous turn is fully complete (all tool calls resolved with outputs).
   */
  async addMessageToConversation(conversationId: string, userMessage: string): Promise<void> {
    const openAI = await this.openAI();
    await openAI.conversations.items.create(conversationId, {
      items: [{ type: 'message', role: 'user', content: userMessage }]
    });
  }

  /**
   * Start the response stream for a conversation (POST openai/responses, stream: true).
   * Use for the first message in a conversation only.
   */
  async startResponseStream(
    conversationId: string,
    agentNameForApi: string,
    signal?: AbortSignal
  ): Promise<Response> {
    const token = await this.getToken();
    const url = `${this._endpoint}/openai/responses?api-version=${CONVERSATIONS_API_VERSION}`;
    const res = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        conversation: conversationId,
        stream: true,
        agent: { name: agentNameForApi, type: 'agent_reference' }
      }),
      signal
    });
    return res;
  }

  /**
   * Single request per turn: the user message is sent as `input` with `conversation` (Foundry appends it to the
   * conversation and streams the answer) — no separate items.create round trip.
   */
  async startResponseStreamWithInput(
    conversationId: string,
    userMessage: string,
    agentNameForApi: string,
    signal?: AbortSignal
  ): Promise<Response> {
    const token = await this.getToken();
    const url = `${this._endpoint}/openai/responses?api-version=${CONVERSATIONS_API_VERSION}`;
    return fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        conversation: conversationId,
        input: [{ type: 'message', role: 'user', content: userMessage }],
        stream: true,
        agent: { name: agentNameForApi, type: 'agent_reference' }
      }),
      signal
    });
  }

  /**
   * Start the response stream by continuing from a previous response (next user turn).
   * Use when the previous turn is fully complete (all tool outputs submitted). Avoids 400
   * "No tool output found for function call" when starting the next message.
   */
  async startResponseStreamWithContinuation(
    previousResponseId: string,
    userMessage: string,
    agentNameForApi: string,
    signal?: AbortSignal
  ): Promise<Response> {
    const token = await this.getToken();
    const url = `${this._endpoint}/openai/responses?api-version=${CONVERSATIONS_API_VERSION}`;
    const res = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        previous_response_id: previousResponseId,
        input: [{ type: 'message', role: 'user', content: userMessage }],
        stream: true,
        agent: { name: agentNameForApi, type: 'agent_reference' }
      }),
      signal
    });
    return res;
  }

  /**
   * Submit tool outputs for a response and get the next response (continuation).
   * Ensures the turn is completed on the server so the next user message does not get 400.
   */
  async submitToolOutputsAndGetNextResponse(
    agentNameForApi: string,
    previousResponseId: string,
    toolOutputs: Array<{ callId: string; output: string }>
  ): Promise<IResponsesApiResponse> {
    const token = await this.getToken();
    const input = toolOutputs.map(({ callId, output }) => ({
      type: 'function_call_output',
      call_id: callId,
      output
    }));
    return createResponse(this._endpoint, token, {
      agent: { name: agentNameForApi, type: 'agent_reference' },
      previous_response_id: previousResponseId,
      input
    });
  }

  /**
   * Conversation history of an agent for the signed-in user.
   *
   * Foundry's agent_name filter only returns conversations in which an answer of that agent completed: a conversation
   * whose first answer failed (sign-in requested, error) or that was continued in Word (agent "<name>-word") was missing.
   * So the full list is read and each conversation is kept when metadata.agent names this agent, or Foundry links it
   * to the agent (or its Word variant). Conversations of another owner are left out, empty ones are hidden, and a
   * missing title is rebuilt once from the first message (then saved in metadata).
   */
  async listConversations(agentName?: string, options?: ListConversationsOptions): Promise<ListConversationsResult> {
    const limit = Math.min(100, Math.max(1, options?.limit ?? 20));
    const base = (agentName ?? '').trim().replace(/-word$/, '');
    const owner = options?.ownerUpn?.trim().toLowerCase();
    try {
      const token = await this._getTokenFn(AZURE_AI_SCOPE);
      const headers = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
      const page = async (agent?: string, after?: string): Promise<{ list: FoundryConversationItem[]; hasMore: boolean }> => {
        const url = new URL(`${this._endpoint}/openai/conversations`);
        url.searchParams.set('api-version', CONVERSATIONS_API_VERSION);
        url.searchParams.set('limit', String(agent ? 100 : limit));
        url.searchParams.set('order', 'desc');
        if (agent) url.searchParams.set('agent_name', agent);
        if (after) url.searchParams.set('after', after);
        const res = await fetch(url.toString(), { headers });
        if (!res.ok) throw new Error(`Historique indisponible (${res.status})`);
        const body = (await res.json()) as { data?: FoundryConversationItem[]; has_more?: boolean } | FoundryConversationItem[];
        const list = Array.isArray(body) ? body : (body.data ?? []);
        return { list, hasMore: !Array.isArray(body) && !!body.has_more };
      };

      const [all, linked, linkedWord] = await Promise.all([
        page(undefined, options?.after),
        base && !options?.after ? page(base).catch(() => ({ list: [], hasMore: false })) : Promise.resolve({ list: [], hasMore: false }),
        base && !options?.after ? page(`${base}-word`).catch(() => ({ list: [], hasMore: false })) : Promise.resolve({ list: [], hasMore: false })
      ]);
      const linkedIds = new Set([...linked.list, ...linkedWord.list].map((c) => c.id));
      const onlyWithTitle = options?.onlyWithTitleInMeta === true;

      const items: IConversationSummary[] = [];
      const toCheck: IConversationSummary[] = [];
      for (const c of all.list) {
        const meta = (c.metadata ?? {}) as Record<string, unknown>;
        const metaAgent = typeof meta.agent === 'string' ? meta.agent : '';
        const metaOwner = typeof meta.owner === 'string' ? meta.owner.toLowerCase() : '';
        if (base && !(metaAgent === base || linkedIds.has(c.id))) continue;
        if (owner && metaOwner && metaOwner !== owner) continue;
        if (meta.hidden === '1') continue;
        const titleStr = (typeof meta.title === 'string' ? meta.title : typeof meta.first_prompt === 'string' ? meta.first_prompt : '').trim();
        if (onlyWithTitle && !titleStr && !metaAgent && !linkedIds.has(c.id)) continue;
        const item: IConversationSummary = {
          id: c.id,
          title: titleStr || 'Conversation',
          updatedAt: (c.created_at !== undefined && c.created_at !== null) ? new Date(c.created_at * 1000) : undefined,
          creationSiteUrl: typeof meta.creation_site_url === 'string' ? meta.creation_site_url : undefined,
          lastUsedSiteUrl: typeof meta.last_used_site_url === 'string' ? meta.last_used_site_url : undefined,
          lastItemId: typeof meta.last_item_id === 'string' ? meta.last_item_id : undefined,
          lastSeenItemId: typeof meta.last_seen_item_id === 'string' ? meta.last_seen_item_id : undefined,
          rawMeta: meta
        };
        // No title, or never answered (no item id recorded and not linked by Foundry): check the content once.
        if (!titleStr || (!item.lastItemId && !item.lastSeenItemId && !linkedIds.has(c.id))) toCheck.push(item);
        items.push(item);
      }

      // Rebuild missing titles from the first message and hide conversations that stayed empty.
      const hidden = new Set<string>();
      await Promise.all(toCheck.slice(0, 15).map(async (item) => {
        try {
          const url = new URL(`${this._endpoint}/openai/conversations/${item.id}/items`);
          url.searchParams.set('api-version', CONVERSATIONS_API_VERSION);
          url.searchParams.set('limit', '20');
          url.searchParams.set('order', 'asc');
          const res = await fetch(url.toString(), { headers });
          if (!res.ok) return;
          const data = ((await res.json()) as { data?: Array<{ type?: string; role?: string; content?: Array<{ text?: string }> }> }).data ?? [];
          if (!data.length) {
            // Created a while ago and still empty: the first answer failed before anything was saved.
            if (item.updatedAt && Date.now() - item.updatedAt.getTime() > 10 * 60_000) hidden.add(item.id);
            return;
          }
          if (item.title !== 'Conversation') return;
          const first = data.find((i) => i.type === 'message' && i.role === 'user');
          const text = cropTitle(stripContextPrefix(first?.content?.map((p) => p.text ?? '').join(' ') ?? ''));
          if (!text) return;
          item.title = text;
          this.updateConversationMetadata(item.id, { title: text }).catch(() => undefined);
        } catch {
          /* keep the item as is */
        }
      }));
      for (const item of items) {
        if (item.title === 'Conversation' && item.updatedAt) {
          item.title = `Conversation du ${item.updatedAt.toLocaleDateString('fr-CH')}`;
        }
      }

      const nextCursor = all.hasMore && all.list.length > 0 ? all.list[all.list.length - 1].id : undefined;
      return { items: items.filter((i) => !hidden.has(i.id)), nextCursor };
    } catch {
      return { items: [] };
    }
  }

  /**
   * Get a single conversation by id (for metadata merge before update).
   */
  async getConversation(conversationId: string): Promise<{ id: string; metadata?: Record<string, unknown> }> {
    const token = await this._getTokenFn(AZURE_AI_SCOPE);
    const url = new URL(`${this._endpoint}/openai/conversations/${conversationId}`);
    url.searchParams.set('api-version', CONVERSATIONS_API_VERSION);
    const res = await fetch(url.toString(), {
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }
    });
    if (!res.ok) throw new Error(`Failed to get conversation: ${res.status}`);
    const body = (await res.json()) as { id?: string; metadata?: Record<string, unknown> };
    return { id: body.id ?? conversationId, metadata: body.metadata };
  }

  /**
   * Update conversation metadata. Merges patch with existing metadata (fetched if not provided).
   */
  async updateConversationMetadata(
    conversationId: string,
    patch: Record<string, unknown>,
    existingMetadata?: Record<string, unknown>
  ): Promise<void> {
    let meta: Record<string, unknown>;
    if (existingMetadata !== undefined) {
      meta = { ...existingMetadata, ...patch };
    } else {
      const conv = await this.getConversation(conversationId);
      meta = { ...(conv.metadata ?? {}), ...patch };
    }
    const openAI = await this.openAI();
    await openAI.conversations.update(conversationId, { metadata: metadataForUpdate(meta) });
  }

  /**
   * Delete a conversation by id via the Foundry REST endpoint.
   * Resolves silently when the conversation is already gone (404).
   */
  async deleteConversation(conversationId: string): Promise<void> {
    const token = await this._getTokenFn(AZURE_AI_SCOPE);
    const url = new URL(`${this._endpoint}/openai/conversations/${conversationId}`);
    url.searchParams.set('api-version', CONVERSATIONS_API_VERSION);
    const res = await fetch(url.toString(), {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok && res.status !== 404) {
      throw new Error(`Failed to delete conversation: ${res.status}`);
    }
  }

  /**
   * Load the latest N messages of a conversation using cursor-based pagination.
   * Returns messages in ascending order (oldest first) suitable for display.
   * Pass `after` (the `nextCursor` from a previous call) to load the page of
   * messages that are older than the current set.
   */
  async getConversationItemsPaged(
    conversationId: string,
    options?: { limit?: number; after?: string }
  ): Promise<{ messages: IChatMessage[]; hasMore: boolean; nextCursor?: string }> {
    const limit = Math.min(50, Math.max(1, options?.limit ?? 20));
    try {
      const token = await this._getTokenFn(AZURE_AI_SCOPE);
      const url = new URL(`${this._endpoint}/openai/conversations/${conversationId}/items`);
      url.searchParams.set('api-version', CONVERSATIONS_API_VERSION);
      url.searchParams.set('limit', String(limit));
      url.searchParams.set('order', 'desc'); // newest first – we reverse before returning
      if (options?.after) url.searchParams.set('after', options.after);

      const res = await fetch(url.toString(), {
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }
      });
      if (res.status === 404) throw new ConversationGoneError(conversationId);
      if (!res.ok) return { messages: [], hasMore: false };

      type RawItem = {
        id?: string;
        type?: string;
        role?: string;
        content?: string | Array<{ type?: string; text?: string }>;
      };

      const body = (await res.json()) as
        | { data?: RawItem[]; has_more?: boolean }
        | RawItem[];

      const list = Array.isArray(body) ? body : (body.data ?? []);
      const hasMore =
        (!Array.isArray(body) && (body as { has_more?: boolean }).has_more === true) ||
        list.filter(i => i.type === 'message' && i.role).length >= limit;

      const validItems = list.filter(i => i.type === 'message' && i.role);

      // Oldest first; reasoning and MCP calls between a user message and the answer become its steps.
      const messages: IChatMessage[] = [];
      let pending: AgentStep[] = [];
      for (const raw of [...list].reverse() as Array<RawItem & Record<string, unknown>>) {
        if (raw.type === 'reasoning' && raw.id) {
          const sum = raw.summary;
          const text = typeof sum === 'string' ? sum : Array.isArray(sum) ? (sum as Array<{ text?: string }>).map((x) => x?.text ?? '').filter(Boolean).join('\n\n') : '';
          if (text.trim()) pending.push({ kind: 'thinking', id: String(raw.id), text });
          continue;
        }
        if (raw.type === 'mcp_call' && raw.id) {
          const err = raw.error ? (typeof raw.error === 'string' ? raw.error : JSON.stringify(raw.error)) : undefined;
          const output = typeof raw.output === 'string' ? raw.output : undefined;
          pending.push({
            kind: 'tool', id: String(raw.id), server: String(raw.server_label ?? ''), name: String(raw.name ?? ''),
            status: err || /^ERREUR/.test(output ?? '') ? 'error' : 'done',
            arguments: typeof raw.arguments === 'string' ? raw.arguments : undefined, output, error: err
          });
          continue;
        }
        if (raw.type !== 'message' || !raw.role) continue;
        const text =
          typeof raw.content === 'string'
            ? raw.content
            : Array.isArray(raw.content)
              ? raw.content.map(c => c?.text ?? '').join('')
              : '';
        if (raw.role === 'user') {
          pending = [];
          messages.push({ role: 'user', id: raw.id, content: stripContextPrefix(text) });
        } else {
          messages.push({ role: 'assistant', id: raw.id, content: text, steps: pending.length ? pending : undefined });
          pending = [];
        }
      }

      // nextCursor = ID of the oldest item in this batch (last in desc list) for the next call
      const nextCursor = validItems.length > 0 ? validItems[validItems.length - 1].id : undefined;

      return { messages, hasMore, nextCursor };
    } catch (err) {
      if (isConversationGone(err)) throw err;
      return { messages: [], hasMore: false };
    }
  }

  /**
   * Fetch the full detail of a single conversation item
   * (GET /openai/conversations/{conversationId}/items/{itemId}).
   */
  async getConversationItem(conversationId: string, itemId: string): Promise<IConversationItemDetail> {
    const token = await this._getTokenFn(AZURE_AI_SCOPE);
    const url = new URL(`${this._endpoint}/openai/conversations/${conversationId}/items/${itemId}`);
    url.searchParams.set('api-version', CONVERSATIONS_API_VERSION);
    const res = await fetch(url.toString(), {
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }
    });
    if (!res.ok) throw new Error(`Failed to get conversation item: ${res.status}`);
    return (await res.json()) as IConversationItemDetail;
  }

  /**
   * Fetch full response metadata by response ID
   * (GET /openai/responses/{responseId}).
   * Includes model name, output items (reasoning, tool calls), and detailed usage.
   */
  async getResponseDetail(responseId: string): Promise<IResponseDetail> {
    const token = await this._getTokenFn(AZURE_AI_SCOPE);
    const url = new URL(`${this._endpoint}/openai/responses/${responseId}`);
    url.searchParams.set('api-version', CONVERSATIONS_API_VERSION);
    const res = await fetch(url.toString(), {
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }
    });
    if (!res.ok) throw new Error(`Failed to get response detail: ${res.status}`);
    return (await res.json()) as IResponseDetail;
  }

  /**
   * Fetch combined message details: the conversation item plus its associated
   * response (model, reasoning, tool calls, detailed token usage).
   * Called by the message details panel when the user clicks the "⋯" button.
   */
  async getMessageDetails(conversationId: string, itemId: string): Promise<IMessageDetailsData> {
    const item = await this.getConversationItem(conversationId, itemId);
    let response: IResponseDetail | undefined;
    const responseId = item.created_by?.response_id;
    if (responseId) {
      try {
        response = await this.getResponseDetail(responseId);
      } catch {
        // Response details are optional – display what we have from the item.
      }
    }
    return { item, response };
  }
}
