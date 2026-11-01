/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../../FoundryAgentsWebPartStrings.d.ts" />
/**
 * Live view of what the agent is doing (inspired by Claude / ChatGPT "thinking" and tool-use blocks):
 * an animated status line with elapsed time, the streamed reasoning summary, and each tool action with
 * its key parameter, duration and status (expandable to see parameters and result). Once the answer is
 * written, the block collapses to "Thought and N actions · 12 s".
 */
import * as React from 'react';
import { makeStyles, mergeClasses, shorthands, tokens, Spinner } from '@fluentui/react-components';
import {
  ChevronRight16Regular,
  CheckmarkCircle16Filled,
  ErrorCircle16Filled,
  Lightbulb16Regular,
  DocumentText16Regular,
  Mail16Regular,
  Search16Regular,
  PlugConnected16Regular,
  Code16Regular
} from '@fluentui/react-icons';
import type { AgentStep } from '../../../types/chat';
import { friendlyToolLabel } from '../../../services/mcpTypes';
import strings from 'FoundryAgentsWebPartStrings';

const shimmer = {
  '0%': { backgroundPosition: '-200% 0' },
  '100%': { backgroundPosition: '200% 0' }
};

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', marginBottom: tokens.spacingVerticalS, maxWidth: '760px' },
  header: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    alignSelf: 'flex-start',
    ...shorthands.padding('2px', '6px', '2px', '2px'),
    ...shorthands.border('none'),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: 'transparent',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    cursor: 'pointer',
    ':hover': { color: tokens.colorNeutralForeground1, backgroundColor: tokens.colorSubtleBackgroundHover }
  },
  chevron: { display: 'inline-flex', transitionProperty: 'transform', transitionDuration: '150ms' },
  chevronOpen: { transform: 'rotate(90deg)' },
  live: {
    fontWeight: tokens.fontWeightSemibold,
    backgroundImage: `linear-gradient(90deg, ${tokens.colorNeutralForeground3} 0%, ${tokens.colorNeutralForeground3} 35%, ${tokens.colorBrandForeground1} 50%, ${tokens.colorNeutralForeground3} 65%, ${tokens.colorNeutralForeground3} 100%)`,
    backgroundSize: '200% 100%',
    WebkitBackgroundClip: 'text',
    backgroundClip: 'text',
    color: 'transparent',
    animationName: shimmer,
    animationDuration: '2.2s',
    animationIterationCount: 'infinite',
    animationTimingFunction: 'linear'
  },
  timer: { fontVariantNumeric: 'tabular-nums', color: tokens.colorNeutralForeground4 },
  timeline: {
    display: 'flex',
    flexDirection: 'column',
    marginTop: tokens.spacingVerticalXS,
    marginLeft: '9px',
    ...shorthands.borderLeft('1px', 'solid', tokens.colorNeutralStroke2),
    paddingLeft: tokens.spacingHorizontalM
  },
  row: { display: 'flex', flexDirection: 'column', ...shorthands.padding('3px', '0') },
  rowHead: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    minWidth: 0,
    ...shorthands.border('none'),
    backgroundColor: 'transparent',
    textAlign: 'left',
    ...shorthands.padding('0'),
    cursor: 'pointer',
    ':hover': { color: tokens.colorNeutralForeground1 }
  },
  icon: { flexShrink: 0, color: tokens.colorNeutralForeground3, display: 'inline-flex' },
  label: { fontWeight: tokens.fontWeightSemibold, flexShrink: 0 },
  detail: { color: tokens.colorNeutralForeground3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', minWidth: 0 },
  dur: { marginLeft: 'auto', paddingLeft: tokens.spacingHorizontalS, color: tokens.colorNeutralForeground4, fontVariantNumeric: 'tabular-nums', flexShrink: 0 },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  ko: { color: tokens.colorPaletteRedForeground1 },
  thinking: {
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'pre-wrap',
    ...shorthands.padding('2px', '0', '4px', '24px')
  },
  thinkingClamp: {
    display: '-webkit-box',
    WebkitLineClamp: 3,
    WebkitBoxOrient: 'vertical',
    ...shorthands.overflow('hidden')
  },
  pre: {
    marginTop: '4px',
    marginLeft: '24px',
    maxHeight: '220px',
    ...shorthands.overflow('auto'),
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
    lineHeight: tokens.lineHeightBase200,
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    ...shorthands.padding('6px', '8px')
  },
  preTitle: { fontWeight: tokens.fontWeightSemibold, display: 'block', marginBottom: '2px' }
});

function serverName(server: string): string {
  if (server === 'ms365') return 'Microsoft 365';
  if (server === 'client') return strings.ActivityServerBrowser;
  return server;
}

function parseArgs(raw?: string): Record<string, unknown> {
  if (!raw) return {};
  try {
    const v = JSON.parse(raw) as unknown;
    return v && typeof v === 'object' ? (v as Record<string, unknown>) : {};
  } catch {
    return {};
  }
}

/** Short human detail for a tool call (the parameter that matters). */
function toolDetail(name: string, rawArgs?: string): string {
  const a = parseArgs(rawArgs);
  const str = (v: unknown): string => (typeof v === 'string' ? v : typeof v === 'number' ? String(v) : '');
  if (name === 'call_tool') return str(a.name).replace(/-/g, ' ');
  if (name === 'get_tool_schema') return Array.isArray(a.names) ? (a.names as string[]).map((n) => n.replace(/-/g, ' ')).join(', ') : '';
  const direct = str(a.query) || str(a.fileName) || str(a.name);
  if (direct) return strings.QuotedText.replace('{0}', direct);
  if (a.searchFields && typeof a.searchFields === 'object') {
    const vals = Object.values(a.searchFields as Record<string, unknown>).map(str).filter(Boolean);
    if (vals.length) return strings.QuotedText.replace('{0}', vals.join(', '));
  }
  if (Array.isArray(a.operations)) return strings.ActivityOperationsCount.replace('{0}', String(a.operations.length));
  return '';
}

function toolIcon(server: string, name: string): React.ReactElement {
  if (/^doc_/.test(name)) return <DocumentText16Regular />;
  if (name === 'search_tools' || name === 'get_tool_schema') return <Search16Regular />;
  if (server === 'ms365') return /mail|message/i.test(name) ? <Mail16Regular /> : <PlugConnected16Regular />;
  return <Code16Regular />;
}

function fmtSeconds(ms: number): string {
  const s = Math.max(0, ms) / 1000;
  return strings.ActivitySeconds.replace('{0}', s < 10 ? s.toFixed(1).replace('.', strings.DecimalSeparator) : String(Math.round(s)));
}

function pretty(raw?: string, max = 1800): string {
  if (!raw) return '';
  let out = raw;
  try { out = JSON.stringify(JSON.parse(raw), null, 2); } catch { /* raw text */ }
  return out.length > max ? `${out.slice(0, max)}\n${strings.ActivityMoreCharacters.replace('{0}', String(out.length - max))}` : out;
}

export interface IAgentActivityProps {
  steps: AgentStep[];
  /** The turn is still running. */
  streaming: boolean;
  /** The answer text has started (the block collapses by default). */
  hasAnswer: boolean;
  startedAt?: number;
  endedAt?: number;
}

const ToolRow: React.FC<{ step: Extract<AgentStep, { kind: 'tool' }> }> = ({ step }) => {
  const s = useStyles();
  const [open, setOpen] = React.useState(false);
  const detail = toolDetail(step.name, step.arguments);
  const dur = step.startedAt && step.endedAt ? fmtSeconds(step.endedAt - step.startedAt) : '';
  const server = serverName(step.server);
  return (
    <div className={s.row}>
      <button type="button" className={s.rowHead} onClick={() => setOpen((o) => !o)} aria-expanded={open}>
        <span className={s.icon}>
          {step.status === 'running' ? <Spinner size="extra-tiny" /> : step.status === 'error' ? <ErrorCircle16Filled className={s.ko} /> : toolIcon(step.server, step.name)}
        </span>
        <span className={s.label}>{friendlyToolLabel(step.name)}</span>
        {detail && <span className={s.detail} title={detail}>{detail}</span>}
        <span className={s.dur}>{server}{dur ? ` · ${dur}` : ''}</span>
        {step.status === 'done' && <CheckmarkCircle16Filled className={s.ok} />}
      </button>
      {open && (
        <>
          {step.arguments && step.arguments !== '{}' && (
            <pre className={s.pre}><span className={s.preTitle}>{strings.ActivityParameters}</span>{pretty(step.arguments, 1200)}</pre>
          )}
          {(step.output || step.error) && (
            <pre className={s.pre}><span className={s.preTitle}>{step.error ? strings.ActivityError : strings.ActivityResult}</span>{pretty(step.error ?? step.output)}</pre>
          )}
        </>
      )}
    </div>
  );
};

const ThinkingRow: React.FC<{ step: Extract<AgentStep, { kind: 'thinking' }>; live: boolean }> = ({ step, live }) => {
  const s = useStyles();
  const [open, setOpen] = React.useState(false);
  if (!step.text.trim()) return null;
  const dur = step.startedAt && step.endedAt ? fmtSeconds(step.endedAt - step.startedAt) : '';
  const clean = step.text.replace(/\*\*/g, '');
  return (
    <div className={s.row}>
      <button type="button" className={s.rowHead} onClick={() => setOpen((o) => !o)} aria-expanded={open || live}>
        <span className={s.icon}><Lightbulb16Regular /></span>
        <span className={s.label}>{strings.ActivityThinking}</span>
        <span className={s.dur}>{dur}</span>
      </button>
      <div className={mergeClasses(s.thinking, !open && !live && s.thinkingClamp)}>{clean}</div>
    </div>
  );
};

export const AgentActivity: React.FC<IAgentActivityProps> = ({ steps, streaming, hasAnswer, startedAt, endedAt }) => {
  const s = useStyles();
  const [open, setOpen] = React.useState<boolean | undefined>(undefined);
  const [now, setNow] = React.useState(Date.now());

  React.useEffect(() => {
    if (!streaming) return;
    const t = window.setInterval(() => setNow(Date.now()), 500);
    return () => window.clearInterval(t);
  }, [streaming]);

  const visible = steps.filter((st) => st.kind === 'tool' || st.text.trim());
  const tools = steps.filter((st) => st.kind === 'tool').length;
  if (!streaming && !visible.length) return null;

  // Expanded while working, collapsed once the answer is being written (unless the user opened it).
  const expanded = open ?? (streaming && !hasAnswer);
  const elapsed = startedAt ? (endedAt ?? (streaming ? now : startedAt)) - startedAt : 0;
  const running = [...steps].reverse().find((st) => st.kind === 'tool' && st.status === 'running');
  const lastThinking = [...steps].reverse().find((st) => st.kind === 'thinking');
  const liveLabel = running && running.kind === 'tool'
    ? friendlyToolLabel(running.name)
    : hasAnswer ? strings.ActivityWritingAnswer : lastThinking ? strings.ActivityThinking : strings.ActivityAnalyzingRequest;
  const summary = tools
    ? (tools > 1 ? strings.ActivitySummaryActions.replace('{0}', String(tools)) : strings.ActivitySummaryOneAction)
    : strings.ActivityThinking;

  return (
    <div className={s.root}>
      <button type="button" className={s.header} onClick={() => setOpen(!expanded)} aria-expanded={expanded}>
        <span className={mergeClasses(s.chevron, expanded && s.chevronOpen)}><ChevronRight16Regular /></span>
        {streaming ? <span className={s.live}>{liveLabel}…</span> : <span>{summary}</span>}
        {elapsed > 0 && <span className={s.timer}>· {fmtSeconds(elapsed)}</span>}
      </button>
      {expanded && visible.length > 0 && (
        <div className={s.timeline}>
          {visible.map((st) =>
            st.kind === 'tool'
              ? <ToolRow key={st.id} step={st} />
              : <ThinkingRow key={st.id} step={st} live={streaming && !st.endedAt} />
          )}
        </div>
      )}
    </div>
  );
};
