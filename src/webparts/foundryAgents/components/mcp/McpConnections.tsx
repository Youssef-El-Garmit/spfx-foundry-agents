/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
/**
 * MCP connections of the current user (Foundry OAuth identity passthrough):
 * header button with an overall status dot, popover listing each server, sign-in buttons,
 * and the in-chat sign-in card shown when Foundry asks for consent during a turn.
 */
import * as React from 'react';
import {
  Button,
  Popover,
  PopoverTrigger,
  PopoverSurface,
  Spinner,
  Tooltip,
  makeStyles,
  tokens,
  mergeClasses
} from '@fluentui/react-components';
import {
  PlugConnected20Regular,
  PlugDisconnected20Regular,
  CheckmarkCircle16Filled,
  Warning16Filled,
  ErrorCircle16Filled,
  QuestionCircle16Regular,
  LockClosed20Regular,
  ArrowSync20Regular
} from '@fluentui/react-icons';
import type { IConsentRequest, IMcpServerState } from '../../services/mcpTypes';
import strings from 'FoundryAgentsWebPartStrings';

const useStyles = makeStyles({
  wrap: { position: 'relative', display: 'inline-flex' },
  dot: {
    position: 'absolute',
    top: '5px',
    right: '5px',
    pointerEvents: 'none',
    width: '8px',
    height: '8px',
    borderRadius: '50%',
    border: `1px solid ${tokens.colorNeutralBackground1}`
  },
  ok: { backgroundColor: tokens.colorPaletteGreenForeground1 },
  warn: { backgroundColor: tokens.colorPaletteDarkOrangeForeground1 },
  bad: { backgroundColor: tokens.colorPaletteRedForeground1 },
  surface: {
    width: '340px',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    backgroundColor: `var(--colorNeutralBackground1, #ffffff)`,
    color: `var(--colorNeutralForeground1, #242424)`,
    boxShadow: tokens.shadow16,
    opacity: 1
  },
  title: { fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase300 },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalXS} 0`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`
  },
  name: { flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 },
  label: { fontWeight: tokens.fontWeightSemibold },
  sub: { fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 },
  note: { fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 },
  footer: { display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: tokens.spacingVerticalS, paddingTop: tokens.spacingVerticalXS },
  card: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalM}`,
    margin: `${tokens.spacingVerticalS} 0 ${tokens.spacingVerticalS} 44px`,
    border: `1px solid ${tokens.colorBrandStroke2}`,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorBrandBackground2,
    maxWidth: '560px'
  },
  cardText: { flex: 1, display: 'flex', flexDirection: 'column', gap: '2px' },
  green: { color: tokens.colorPaletteGreenForeground1 },
  orange: { color: tokens.colorPaletteDarkOrangeForeground1 },
  red: { color: tokens.colorPaletteRedForeground1 }
});

/** Display names of the well-known MCP servers (server_label → display name); other servers show their label. */
export function serverDisplayName(label: string): string {
  if (label === 'ms365') return 'Microsoft 365';
  if (label === 'word') return strings.McpServerWordDocuments;
  return label;
}

/**
 * Open the Foundry consent page in a popup and resolve when the user closes it
 * (Foundry shows "Authentication succeeded" then closes the window itself).
 */
export function openConsentPopup(link: string): Promise<void> {
  return new Promise((resolve) => {
    const w = 520, h = 720;
    const left = window.screenX + Math.max(0, (window.outerWidth - w) / 2);
    const top = window.screenY + Math.max(0, (window.outerHeight - h) / 2);
    const popup = window.open(link, 'foundry-mcp-consent', `width=${w},height=${h},left=${left},top=${top}`);
    if (!popup) {
      window.open(link, '_blank');
      resolve();
      return;
    }
    const timer = window.setInterval(() => {
      if (popup.closed) {
        window.clearInterval(timer);
        resolve();
      }
    }, 500);
  });
}

function StatusIcon({ status }: { status: IMcpServerState['status'] }): React.ReactElement {
  const s = useStyles();
  if (status === 'connected') return <CheckmarkCircle16Filled className={s.green} />;
  if (status === 'consent_required') return <Warning16Filled className={s.orange} />;
  if (status === 'error') return <ErrorCircle16Filled className={s.red} />;
  return <QuestionCircle16Regular />;
}

function statusText(st: IMcpServerState): string {
  switch (st.status) {
    case 'connected': return st.toolCount ? strings.McpStatusConnectedTools.replace('{0}', String(st.toolCount)) : strings.McpStatusConnected;
    case 'consent_required': return strings.McpStatusConsentRequired;
    case 'error': return st.error ? strings.McpStatusErrorDetail.replace('{0}', st.error.slice(0, 80)) : strings.McpStatusError;
    default: return strings.McpStatusUnknown;
  }
}

export interface IMcpConnectionsButtonProps {
  servers: IMcpServerState[];
  checking: boolean;
  onCheck: () => void;
  onSignIn: (server: IMcpServerState) => void;
}

export const McpConnectionsButton: React.FC<IMcpConnectionsButtonProps> = ({ servers, checking, onCheck, onSignIn }) => {
  const s = useStyles();
  const [open, setOpen] = React.useState(false);
  if (!servers.length) return null;
  const needs = servers.filter((x) => x.status === 'consent_required' || x.status === 'error').length;
  const allOk = servers.every((x) => x.status === 'connected');
  const dot = allOk ? s.ok : needs ? s.warn : undefined;
  const tip = allOk ? strings.McpTooltipAllConnected : needs ? strings.McpTooltipNeedsSignIn.replace('{0}', String(needs)) : strings.McpConnectionsTitle;
  return (
    <Popover open={open} onOpenChange={(_, d) => { setOpen(d.open); if (d.open && servers.some((x) => x.status === 'unknown')) onCheck(); }} positioning="below-end" withArrow>
      <PopoverTrigger disableButtonEnhancement>
        <span className={s.wrap}>
          <Tooltip content={tip} relationship="label">
            <Button
              appearance="subtle"
              shape="circular"
              icon={allOk ? <PlugConnected20Regular /> : <PlugDisconnected20Regular />}
              aria-label={tip}
            />
          </Tooltip>
          {dot && <span className={mergeClasses(s.dot, dot)} aria-hidden />}
        </span>
      </PopoverTrigger>
      <PopoverSurface className={s.surface}>
        <span className={s.title}>{strings.McpConnectionsTitle}</span>
        <span className={s.note}>{strings.McpConnectionsNote}</span>
        {servers.map((st) => (
          <div key={st.label} className={s.row}>
            <StatusIcon status={st.status} />
            <div className={s.name}>
              <span className={s.label}>{serverDisplayName(st.label)}</span>
              <span className={s.sub}>{statusText(st)}</span>
            </div>
            {st.status === 'consent_required' && st.consentLink && (
              <Button size="small" appearance="primary" icon={<LockClosed20Regular />} onClick={() => onSignIn(st)}>
                {strings.McpSignIn}
              </Button>
            )}
          </div>
        ))}
        <div className={s.footer}>
          <Button size="small" icon={checking ? <Spinner size="extra-tiny" /> : <ArrowSync20Regular />} disabled={checking} onClick={onCheck}>
            {strings.McpCheckConnections}
          </Button>
        </div>
      </PopoverSurface>
    </Popover>
  );
};

export interface IConsentCardProps {
  consent: IConsentRequest;
  busy?: boolean;
  onSignIn: (consent: IConsentRequest) => void;
}

/** In-chat card: the agent needs the user to sign in to an MCP server to continue. */
export const ConsentCard: React.FC<IConsentCardProps> = ({ consent, busy, onSignIn }) => {
  const s = useStyles();
  const name = serverDisplayName(consent.serverLabel);
  return (
    <div className={s.card} role="group" aria-label={strings.McpConsentAriaLabel.replace('{0}', name)}>
      {consent.done ? <CheckmarkCircle16Filled className={s.green} /> : <LockClosed20Regular />}
      <div className={s.cardText}>
        <span className={s.label}>{consent.done ? strings.McpConsentDone.replace('{0}', name) : strings.McpConsentRequired.replace('{0}', name)}</span>
        <span className={s.sub}>
          {consent.done
            ? strings.McpConsentDoneDetail
            : strings.McpConsentRequiredDetail.replace('{0}', name)}
        </span>
      </div>
      {!consent.done && (
        <Button appearance="primary" icon={busy ? <Spinner size="extra-tiny" /> : <LockClosed20Regular />} disabled={busy} onClick={() => onSignIn(consent)}>
          {strings.McpSignIn}
        </Button>
      )}
    </div>
  );
};
