import * as React from 'react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IAgentConfig } from '../models';
import { AudienceService } from '../services/AudienceService';

export type AudienceFilterStatus = 'loading' | 'ready';

export interface IUseAudienceFilterResult {
  /** Agents the current user is allowed to see (audience-filtered). */
  visibleAgents: IAgentConfig[];
  /** Whether the membership checks are still in flight. */
  status: AudienceFilterStatus;
  /**
   * Re-check whether a specific agent is accessible (e.g. after a direct-link navigation
   * to an agent with audience restrictions). Resolves to `true` if access is granted.
   */
  checkAgent: (agent: IAgentConfig) => Promise<boolean>;
}

/**
 * Filters `agents` by audience rules.
 *
 * - Agents without any `audienceGroups` entry are always visible.
 * - Agents with audience entries are shown only when the current user belongs to at
 *   least one of the listed groups.
 *
 * All membership checks are batched in a single useEffect and the service caches
 * individual group membership responses so subsequent calls (e.g. `checkAgent`) are instant.
 */
export function useAudienceFilter(
  agents: IAgentConfig[],
  context: WebPartContext
): IUseAudienceFilterResult {
  const serviceRef = React.useRef<AudienceService | null>(null);
  if (!serviceRef.current) {
    serviceRef.current = new AudienceService(context);
  }
  const service = serviceRef.current;

  const [visibleAgents, setVisibleAgents] = React.useState<IAgentConfig[]>([]);
  const [status, setStatus] = React.useState<AudienceFilterStatus>('loading');

  React.useEffect(() => {
    let cancelled = false;
    setStatus('loading');

    Promise.all(
      agents.map(async (agent) => {
        const allowed = await service.canAccessAgent(agent.audienceGroups);
        return allowed ? agent : null;
      })
    ).then((results) => {
      if (cancelled) return;
      setVisibleAgents(results.filter((a): a is IAgentConfig => a !== null));
      setStatus('ready');
    }).catch(() => {
      if (cancelled) return;
      // On error, fall back to showing all agents (fail-open) so users are not blocked
      setVisibleAgents(agents);
      setStatus('ready');
    });

    return () => { cancelled = true; };
    // Re-run whenever the set of agents changes (config update) or context changes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [agents, context]);

  const checkAgent = React.useCallback(
    (agent: IAgentConfig): Promise<boolean> => {
      return service.canAccessAgent(agent.audienceGroups);
    },
    [service]
  );

  return { visibleAgents, status, checkAgent };
}
