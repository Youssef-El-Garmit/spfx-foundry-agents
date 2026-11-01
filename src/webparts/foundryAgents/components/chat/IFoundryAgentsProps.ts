import type { IAgentConfig } from '../../models';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFoundryAgentsProps {
  /** List of agents configured in the property pane */
  agents: IAgentConfig[];
  /** Whether the page is in edit mode (e.g. to show "open properties" when not configured) */
  isEditMode?: boolean;
  /** Callback to open the web part property pane */
  onOpenPropertyPane: () => void;
  /** When set (e.g. by host opening properties in Personal App), open custom settings panel */
  openSettingsTrigger?: number;
  /** SPFx context (for AAD token and theme) */
  context: WebPartContext;
  /** CSS value for the web part's minimum height (e.g. "min(600px, 90vh)") */
  minHeight?: string;
}
