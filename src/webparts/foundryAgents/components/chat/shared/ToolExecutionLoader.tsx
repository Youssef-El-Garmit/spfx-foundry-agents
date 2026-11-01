/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../../FoundryAgentsWebPartStrings.d.ts" />
/**
 * Pretty loading indicator for tool execution, using Fluent UI / Copilot-style.
 * Shown while the agent is waiting for a tool result (e.g. Graph/SharePoint API call).
 */

import * as React from 'react';
import { Spinner, Text } from '@fluentui/react-components';
import strings from 'FoundryAgentsWebPartStrings';
import styles from './ToolExecutionLoader.module.scss';

export interface IToolExecutionLoaderProps {
  /** Current tool name being executed (e.g. "microsoft_graph_get") */
  toolName?: string;
  /** Accessible label */
  ariaLabel?: string;
  /** Label shown next to spinner (e.g. "Running tool") */
  runningLabel?: string;
  /** Fallback when no tool name (e.g. "Tool") */
  toolFallback?: string;
  className?: string;
}

export const ToolExecutionLoader: React.FunctionComponent<IToolExecutionLoaderProps> = ({
  toolName,
  ariaLabel = strings.RunningTool,
  runningLabel = strings.RunningTool,
  toolFallback = strings.ToolGeneric,
  className
}) => {
  const displayName = toolName ? formatToolName(toolName) : toolFallback;
  return (
    <div
      className={`${styles.root} ${className ?? ''}`}
      role="status"
      aria-live="polite"
      aria-label={toolName ? `${ariaLabel}: ${displayName}` : ariaLabel}
    >
      <div className={styles.iconWrap}>
        <Spinner size="small" className={styles.spinner} />
      </div>
      <div className={styles.content}>
        <Text className={styles.label}>{runningLabel}</Text>
        <Text className={styles.toolName}>{displayName}</Text>
      </div>
    </div>
  );
};

function formatToolName(name: string): string {
  return name
    .replace(/_/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase())
    .replace(/Microsoft Graph /gi, 'Microsoft Graph ');
}
