import * as React from 'react';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** State parsed from the URL hash. */
export interface IAgentHashRouteState {
  /**
   * Slugified display name from URL (e.g. "hr-assistant").
   * Undefined when no agent segment is present in the hash.
   */
  agentSlug: string | undefined;
  /** Conversation ID from URL (undefined when absent). */
  conversationId: string | undefined;
}

export interface IAgentHashRoute extends IAgentHashRouteState {
  /**
   * Update `window.location.hash` to reflect the given agent (and optionally conversation).
   * This is a no-op when the hash already matches – safe to call on every render.
   *
   * Hash format:
   *   #/agent/{displayNameSlug}                             – no active conversation
   *   #/agent/{displayNameSlug}/conversation/{conversationId} – active conversation
   */
  navigateTo: (displayName: string, conversationId?: string) => void;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers (exported so consumers can match agents by slug without re-importing)
// ─────────────────────────────────────────────────────────────────────────────

const AGENT_SEGMENT = 'agent';

/**
 * Turn a display name into a URL-safe slug.
 * Exported so FoundryAgents can match agents by slugifying their displayName.
 */
export function slugify(text: string): string {
  return (text ?? '')
    .toLowerCase()
    .replace(/\s+/g, '-')
    .replace(/[^\w-]/g, '')
    .replace(/-{2,}/g, '-')
    .slice(0, 60) || AGENT_SEGMENT;
}

/**
 * Build the hash string for an agent + optional conversation.
 *
 * Format:
 *   #/agent/{displayNameSlug}
 *   #/agent/{displayNameSlug}/conversation/{conversationId}
 */
export function buildAgentHash(displayName: string, conversationId?: string): string {
  const slug = slugify(displayName);
  const base = `#/${AGENT_SEGMENT}/${slug}`;
  return conversationId ? `${base}/conversation/${encodeURIComponent(conversationId)}` : base;
}

/** Parse `window.location.hash` into a route state object. */
export function parseAgentHash(hash: string): IAgentHashRouteState {
  const path = hash.startsWith('#') ? hash.slice(1) : hash;
  const prefix = `/${AGENT_SEGMENT}/`;

  if (!path.startsWith(prefix)) {
    return { agentSlug: undefined, conversationId: undefined };
  }

  const parts = path.slice(prefix.length).split('/');
  const agentSlug = parts[0] ? decodeURIComponent(parts[0]) : undefined;

  const convIdx = parts.indexOf('conversation');
  const conversationId =
    convIdx !== -1 && parts[convIdx + 1]
      ? decodeURIComponent(parts[convIdx + 1])
      : undefined;

  return { agentSlug, conversationId };
}

// ─────────────────────────────────────────────────────────────────────────────
// Hook
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Lightweight hash-based routing for FoundryAgents (no React-Router dependency).
 *
 * URL format:
 *   #/agent/{displayNameSlug}                             – agent selected, no conversation
 *   #/agent/{displayNameSlug}/conversation/{conversationId} – agent + conversation selected
 *
 * Agents are identified by their slugified display name.
 * Display names must therefore be unique within a web part instance
 * (enforced at edit time by the AgentEditDialog duplicate-name validation).
 */
export function useAgentHashRoute(): IAgentHashRoute {
  const [routeState, setRouteState] = React.useState<IAgentHashRouteState>(
    () => parseAgentHash(window.location.hash)
  );

  React.useEffect(() => {
    const handleHashChange = (): void => {
      setRouteState(parseAgentHash(window.location.hash));
    };
    window.addEventListener('hashchange', handleHashChange);
    return () => window.removeEventListener('hashchange', handleHashChange);
  }, []);

  const navigateTo = React.useCallback(
    (displayName: string, conversationId?: string): void => {
      const next = buildAgentHash(displayName, conversationId);
      if (window.location.hash !== next) {
        window.location.hash = next;
      }
    },
    []
  );

  return { ...routeState, navigateTo };
}
