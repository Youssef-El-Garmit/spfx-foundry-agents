/**
 * Icon source: Fluent 2 icon name or custom image URL.
 */
export type IconType = 'fluent' | 'custom';

/**
 * Discriminates which directory backend an audience entry comes from.
 * - 'graph'       : Azure AD / Entra group (objectId stored in `id`)
 * - 'sharepoint'  : SharePoint site group (numeric Id stored in `id`)
 * - 'user'        : Individual Entra user (objectId stored in `id`, UPN in `mail`)
 */
export type AudienceGroupType = 'graph' | 'sharepoint' | 'user';

/**
 * A single audience entry – an Entra group, a SharePoint site group, or an individual user.
 */
export interface IAudienceGroup {
  /** objectId for graph/user; numeric Id as string for sharepoint */
  id: string;
  /** Human-readable display name */
  displayName: string;
  /** Entry source */
  type: AudienceGroupType;
  /** Email / UPN – used for display and for user-type equality check */
  mail?: string;
}

/**
 * Single agent configuration (same properties for each agent in the list).
 * Aligns with Azure AI Foundry Agent Service requirements:
 * - endpoint: resource endpoint (e.g. https://<resource>.services.ai.azure.com/api/projects/<project>)
 * - agentId: assistant/agent identifier used in API paths
 */
export interface IAgentConfig {
  /** Display name shown in UI */
  displayName: string;
  /** Azure AI Foundry agent ID (assistant ID) */
  agentId: string;
  /** Endpoint URL (e.g. https://<resource>.services.ai.azure.com/api/projects/<project> or project endpoint) */
  endpointUrl: string;
  /** Icon type: Fluent 2 icon or custom image */
  iconType: IconType;
  /** Fluent 2 icon name (e.g. "Bot", "Chat") when iconType === 'fluent' */
  fluentIconName?: string;
  /** Custom icon image URL when iconType === 'custom' */
  customIconUrl?: string;
  /** Suggestion prompts shown in empty conversation (fallback: web part default list) */
  suggestionPrompts?: string[];
  /** When true, history shows only conversations that have a title in metadata (started from this web part). When false, all conversations for this agent are shown. */
  showOnlyWebPartConversations?: boolean;
  /** When true, the chat input shows an attachment button so users can attach files to messages. */
  allowAttachments?: boolean;
  /**
   * Audience restriction – list of Entra groups or SharePoint site groups that may access this agent.
   * When the array is empty or undefined the agent is visible to everyone.
   * When at least one entry is present the current user must belong to at least one of them.
   */
  audienceGroups?: IAudienceGroup[];
  /**
   * When false (default), function_call from the agent are not executed; the model receives an error instead.
   * When true, function calls are executed subject to blockedHttpMethods.
   */
  allowFunctionCalls?: boolean;
  /**
   * HTTP methods that must not be executed for function_call (e.g. DELETE, PATCH).
   * Default when undefined or empty: ['DELETE', 'PATCH'].
   */
  blockedHttpMethods?: string[];
  /**
   * Optional template for additional context sent to the model (user, datetime, timezone, region).
   * Use variables like {userDisplayName}, {userMail}, {currentDateTime}, {currentTimezone}, {currentRegion}.
   */
  additionalInformationTemplate?: string;
  /**
   * Optional full URL of an Azure OpenAI / Foundry audio transcriptions endpoint used for voice dictation
   * (e.g. https://<resource>.services.ai.azure.com/openai/deployments/<deployment>/audio/transcriptions?api-version=...).
   * Called with the user's Entra token (audience https://ai.azure.com). When empty, the browser speech recognition is used.
   */
  transcriptionUrl?: string;
  /** Optional vocabulary hint sent as the transcription prompt (names and terms to spell correctly). */
  transcriptionPrompt?: string;
}
