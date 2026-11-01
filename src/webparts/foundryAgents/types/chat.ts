/** Citation/annotation from AI agent responses (Azure AI Agent SDK annotation types). */
export interface IAnnotation {
  /** Type: "uri_citation", "file_citation", "file_path", or "container_file_citation" */
  type: 'uri_citation' | 'file_citation' | 'file_path' | 'container_file_citation';
  /** Display label (title or filename) */
  label: string;
  /** URL for URI citations */
  url?: string;
  /** File ID for file citations */
  fileId?: string;
  /** Placeholder text in the response to replace (e.g., "【4:0†source】") */
  textToReplace?: string;
  /** Start index in the text where the citation applies */
  startIndex?: number;
  /** End index in the text where the citation applies */
  endIndex?: number;
  /** Quote from the source document (for file citations) */
  quote?: string;
}

/** Token/usage info from agent response (optional). */
export interface IUsageInfo {
  promptTokens?: number;
  completionTokens?: number;
  totalTokens?: number;
}

/** Optional metadata for assistant messages (timestamp, usage, etc.). */
export interface IAssistantMessageMore {
  time?: string;
  usage?: IUsageInfo;
}

import type { IConsentRequest, IDocumentRef } from '../services/mcpTypes';

/** One step of the agent's work, shown live like Claude / ChatGPT (thinking, then each tool action). */
export type AgentStep =
  | {
      kind: 'thinking';
      id: string;
      /** Reasoning summary streamed by the model (may stay empty). */
      text: string;
      startedAt?: number;
      endedAt?: number;
    }
  | {
      kind: 'tool';
      id: string;
      /** MCP server label ("ms365", "crm") or "client" for function tools. */
      server: string;
      name: string;
      status: 'running' | 'done' | 'error';
      arguments?: string;
      output?: string;
      error?: string;
      startedAt?: number;
      endedAt?: number;
    };

/** Chat message (user or assistant). */
export interface IChatMessage {
  role: 'user' | 'assistant';
  content: string;
  /** Optional stable id (e.g. from API); fallback to index when missing */
  id?: string;
  /** Citations/sources from agent (e.g. from 【N†source】 placeholders or API annotations) */
  annotations?: IAnnotation[];
  /** Optional metadata (timestamp, usage) */
  more?: IAssistantMessageMore;
  /** Response duration in ms (optional) */
  duration?: number;
  /** Files attached to this user message (populated for live-session messages; history relies on content parsing). */
  uploadedFiles?: Array<{ name: string; fileId: string }>;
  /** Sign-in requested by Foundry for an MCP server during this assistant turn. */
  consent?: IConsentRequest;
  /** Word documents created or edited during this assistant turn. */
  documents?: IDocumentRef[];
  /** Agent work for this turn (reasoning summaries and tool actions), in order. */
  steps?: AgentStep[];
  /** Turn timing (epoch ms) for "worked for N s". */
  startedAt?: number;
  endedAt?: number;
}
