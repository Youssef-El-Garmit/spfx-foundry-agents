/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { ExportImportSettings } from '../components/config/ExportImportSettings';
import type { IAgentConfig } from '../models';
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

export interface IExportImportPropertyPaneControlProps {
  agentsJson: string;
  onChange: (newJson: string) => void;
}

/**
 * Renders in the property pane: Export / Import block only (custom group under Appearance).
 */
export const ExportImportPropertyPaneControl: React.FunctionComponent<IExportImportPropertyPaneControlProps> = ({
  agentsJson,
  onChange,
}) => {
  const agents = React.useMemo(() => parseAgents(agentsJson), [agentsJson]);
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

  return (
    <div ref={setContainerRef} style={{ minHeight: 1 }}>
      <ExportImportSettings
        agents={agents}
        onSave={handleSave}
        portalMountNode={portalTarget}
        showSectionStyle={false}
      />
    </div>
  );
};
