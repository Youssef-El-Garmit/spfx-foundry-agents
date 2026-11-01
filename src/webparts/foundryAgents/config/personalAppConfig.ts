/**
 * Personal app config stored in the user's OneDrive (Apps folder).
 * Allows the same agent list to sync across devices for Teams Personal App.
 */

import { graphfi, SPFx } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/files';
import type { IAgentConfig } from '../models';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

const SETTINGS_FILE = 'nemoura.foundryAgentsPersonalApp.settings.json';

export interface IPersonalAppSettings {
  agents: IAgentConfig[];
}

const DEFAULT_SETTINGS: IPersonalAppSettings = {
  agents: []
};

function parseAgents(value: unknown): IAgentConfig[] {
  if (!Array.isArray(value)) return [];
  return value.filter(
    (a): a is IAgentConfig =>
      a &&
      typeof a === 'object' &&
      typeof (a as IAgentConfig).displayName === 'string' &&
      typeof (a as IAgentConfig).agentId === 'string' &&
      typeof (a as IAgentConfig).endpointUrl === 'string'
  );
}

function createGraph(context: WebPartContext): ReturnType<typeof graphfi> {
  return graphfi().using(SPFx(context));
}

async function loadSettings<T extends IPersonalAppSettings>(context: WebPartContext, defaults: T): Promise<T> {
  try {
    const graph = createGraph(context);
    const content = await graph.me.drive.getItemByPath(`Apps/${SETTINGS_FILE}`).getContent();
    const text = await new Response(content).text();
    const parsed = JSON.parse(text) as Partial<IPersonalAppSettings>;
    return { ...defaults, ...parsed } as T;
  } catch {
    return defaults;
  }
}

async function saveSettings(context: WebPartContext, settings: IPersonalAppSettings): Promise<void> {
  const graph = createGraph(context);
  await graph.me.drive.root.upload({
    content: JSON.stringify(settings, null, 2),
    filePathName: `Apps/${SETTINGS_FILE}`,
    contentType: 'application/json; charset=utf-8'
  });
}

/**
 * Load agents from the user's OneDrive settings file.
 * Returns empty array if file doesn't exist or on any error (silent).
 */
export async function loadPersonalAppAgents(context: WebPartContext): Promise<IAgentConfig[]> {
  const loaded = await loadSettings(context, DEFAULT_SETTINGS);
  return parseAgents(loaded.agents);
}

/**
 * Save agents to the user's OneDrive settings file.
 * Throws on error (network, permissions, etc.).
 */
export async function savePersonalAppAgents(context: WebPartContext, agentsJson: string): Promise<void> {
  const agents = parseAgents(JSON.parse(agentsJson || '[]') as IAgentConfig[]);
  await saveSettings(context, { agents });
}
