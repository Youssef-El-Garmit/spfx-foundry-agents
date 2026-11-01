import * as React from 'react';

/** Teams app.getContext() result: app, channel (only in channel/group tab), team. */
export interface ITeamsAppContext {
  app?: { host?: { name?: string }; [key: string]: unknown };
  channel?: { id?: string; [key: string]: unknown };
  team?: unknown;
}

export interface ITeamsHostResult {
  isTeamsTab: boolean;
  isPersonalApp: boolean;
  isTeams: boolean;
}

function detectTeamsFromEnvironment(): boolean {
  if (typeof window === 'undefined') return false;
  const u = window.location?.href?.toLowerCase() ?? '';
  const r = (typeof document !== 'undefined' ? document.referrer : '')?.toLowerCase() ?? '';
  const inIframe = window.self !== window.top;
  return inIframe && (u.includes('teams.microsoft.com') || u.includes('/teams/') || r.includes('teams'));
}

// eslint-disable-next-line @rushstack/no-new-null
function fromContext(ctx: ITeamsAppContext | undefined | null): ITeamsHostResult {
  if (!ctx || typeof ctx !== 'object') {
    return { isTeamsTab: false, isPersonalApp: false, isTeams: false };
  }
  const hasApp = !!(ctx.app && typeof ctx.app === 'object');
  const hostName = (ctx.app as { host?: { name?: string } } | undefined)?.host?.name;
  const isTeamsFromHost = String(hostName).toLowerCase() === 'teams';
  const hasChannel = !!(ctx.channel && typeof ctx.channel === 'object');
  const hasTeam = !!(ctx.team && typeof ctx.team === 'object');
  const isTeams = isTeamsFromHost || hasApp || hasChannel || hasTeam;
  const isTeamsTab = isTeams && hasChannel;
  const isPersonalApp = isTeams && !hasChannel;
  return { isTeamsTab, isPersonalApp, isTeams };
}

/** Shared with web part: derive Personal App vs channel tab from getContext() result. */
export function getTeamsHostFromContext(ctx: ITeamsAppContext | undefined): { isPersonalApp: boolean; isTeamsTab: boolean } {
  const r = fromContext(ctx);
  return { isPersonalApp: r.isPersonalApp, isTeamsTab: r.isTeamsTab };
}

/**
 * Detects Teams host and Personal App vs channel/group tab via app.getContext().
 * Falls back to environment heuristic if SDK is unavailable.
 */
export function useTeamsHost(context: { sdks?: unknown; sdksAsync?: Promise<unknown> }): ITeamsHostResult {
  const [result, setResult] = React.useState<ITeamsHostResult>(() => ({
    isTeamsTab: false,
    isPersonalApp: false,
    isTeams: detectTeamsFromEnvironment()
  }));
  const mounted = React.useRef(true);

  React.useEffect(() => {
    mounted.current = true;
    let fallbackTimer: number | undefined;
    const resolve = (ctx: ITeamsAppContext | undefined): void => {
      if (!mounted.current) return;
      if (fallbackTimer !== undefined && fallbackTimer !== null) clearTimeout(fallbackTimer);
      setResult(fromContext(ctx));
    };
    const applyFallback = (): void => {
      if (!mounted.current) return;
      setResult((prev) => (prev.isTeams ? prev : { ...prev, isTeams: detectTeamsFromEnvironment() }));
    };
    if (typeof window !== 'undefined' && window.setTimeout) {
      fallbackTimer = window.setTimeout(applyFallback, 2000);
    }
    const runGetContext = (app: { getContext(): Promise<ITeamsAppContext> }): void => {
      app.getContext()
        .then((ctx) => { if (mounted.current) resolve(ctx as ITeamsAppContext); })
        .catch(() => { if (mounted.current) applyFallback(); });
    };

    const sdks = context.sdks as {
      microsoftTeams?: { teamsJs?: { app?: { getContext(): Promise<ITeamsAppContext> } } };
    } | undefined;
    const teamsJs = sdks?.microsoftTeams?.teamsJs;
    const cleanup = (): void => {
      mounted.current = false;
      if (fallbackTimer !== undefined && fallbackTimer !== null) clearTimeout(fallbackTimer);
    };
    if (teamsJs?.app) {
      runGetContext(teamsJs.app);
      return cleanup;
    }
    const sdksAsync = (context as { sdksAsync?: Promise<unknown> }).sdksAsync;
    if (sdksAsync) {
      sdksAsync.then((resolved) => {
        const teams = (resolved as { microsoftTeams?: { teamsJs?: { app?: { getContext(): Promise<ITeamsAppContext> } } } })?.microsoftTeams;
        if (teams?.teamsJs?.app && mounted.current) {
          runGetContext(teams.teamsJs.app);
        } else if (mounted.current) {
          applyFallback();
        }
      }).catch(() => { if (mounted.current) applyFallback(); });
      return cleanup;
    }
    applyFallback();
    return cleanup;
  }, [context]);

  return result;
}
