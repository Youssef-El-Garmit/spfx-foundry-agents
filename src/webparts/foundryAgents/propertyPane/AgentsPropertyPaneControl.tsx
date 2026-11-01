/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { AgentConfigPanel } from '../components/config';
import type { IAgentConfig } from '../models';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

function parseAgents(agentsJson: string | undefined): IAgentConfig[] {
  if (!agentsJson || !agentsJson.trim()) return [];
  try {
    const arr = JSON.parse(agentsJson) as unknown;
    if (!Array.isArray(arr)) return [];
    return arr.filter(
      (a): a is IAgentConfig =>
        a &&
        typeof a === 'object' &&
        typeof (a as IAgentConfig).displayName === 'string' &&
        typeof (a as IAgentConfig).agentId === 'string' &&
        typeof (a as IAgentConfig).endpointUrl === 'string'
    );
  } catch {
    return [];
  }
}

export interface IAgentsPropertyPaneControlProps {
  agentsJson: string;
  onChange: (newJson: string) => void;
  /** SPFx context forwarded to AgentEditDialog for audience group search/validation */
  context: WebPartContext;
}

/**
 * Renders in the native property pane: list of agents with add / edit / remove.
 * Uses a portal target inside FluentProvider so Dialog and Tooltip inherit theme (font, colors, borders).
 */
export const AgentsPropertyPaneControl: React.FunctionComponent<IAgentsPropertyPaneControlProps> = ({
  agentsJson,
  onChange,
  context
}) => {
  const initialAgents = React.useMemo(() => parseAgents(agentsJson), [agentsJson]);
  const [portalTarget, setPortalTarget] = React.useState<HTMLElement | undefined>(undefined);
  const setContainerRef = React.useCallback((el: HTMLDivElement | null) => {
    setPortalTarget(el ?? undefined);
  }, []);

  const handleSave = React.useCallback(
    (newAgents: IAgentConfig[]) => {
      const valid = newAgents.filter(
        (a) =>
          (a.displayName ?? '').trim() &&
          (a.agentId ?? '').trim() &&
          (a.endpointUrl ?? '').trim()
      );
      onChange(JSON.stringify(valid.length > 0 ? valid : newAgents));
    },
    [onChange]
  );

  // Workaround for SharePoint FluentProvider ID collision (sp-dev-docs #9847): isolate our
  // FluentProvider with a unique ID prefix so generated IDs don't clash with SP's FluentUI.

  return (
        <div ref={setContainerRef} style={{ minHeight: 1 }}>
          <AgentConfigPanel
            agents={initialAgents}
            onSave={handleSave}
            portalMountNode={portalTarget}
            context={context}
          />
        </div>
  );
};
