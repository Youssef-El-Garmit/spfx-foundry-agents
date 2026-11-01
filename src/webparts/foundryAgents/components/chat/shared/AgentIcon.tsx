import * as React from 'react';
import * as FluentIcons from '@fluentui/react-icons';
import { Bot24Regular } from '@fluentui/react-icons';
import type { IAgentConfig } from '../../../models';

type IconComponent = React.ComponentType<{ className?: string; style?: React.CSSProperties }>;

// Fluent UI icons are React.forwardRef objects: { $$typeof, render }
// They are NOT plain functions, so we detect them by checking for a `render` method.
const EXCLUDED = new Set(['wrapIcon', 'bundleIcon', 'createFluentIcon', 'useIconState', 'useIconContext', 'IconDirectionContextProvider']);

const ICON_MAP: Record<string, IconComponent> = {};
for (const [key, val] of Object.entries(FluentIcons as Record<string, unknown>)) {
  if (EXCLUDED.has(key)) continue;
  if (val && typeof val === 'object' && typeof (val as { render?: unknown }).render === 'function') {
    ICON_MAP[key] = val as IconComponent;
  }
}

/** All available Fluent UI icon names, sorted alphabetically. */
export const ICON_NAMES: string[] = Object.keys(ICON_MAP).sort();

/** Resolve a Fluent UI icon name to its component; falls back to Bot24Regular. */
export function getFluentIcon(name: string): IconComponent {
  return ICON_MAP[(name || '').trim()] ?? Bot24Regular;
}

export interface IAgentIconProps {
  agent: IAgentConfig;
  className?: string;
  style?: React.CSSProperties;
  size?: number;
}

export const AgentIcon: React.FunctionComponent<IAgentIconProps> = ({ agent, className, style, size = 24 }) => {
  if (agent.iconType === 'custom' && agent.customIconUrl) {
    return (
      <img
        src={agent.customIconUrl}
        alt=""
        className={className}
        style={{ width: `${size}px`, height: `${size}px`, borderRadius: '50%', objectFit: 'cover', ...style }}
      />
    );
  }
  const Icon = getFluentIcon(agent.fluentIconName || 'Bot24Regular');
  return <Icon className={className} style={{ width: `${size}px`, height: `${size}px`, ...style }} />;
};
