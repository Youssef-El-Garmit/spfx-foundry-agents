/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import {
  FluentProvider,
  IdPrefixProvider,
  PortalMountNodeProvider,
  Toaster,
  useToastController,
  Toast,
  ToastTitle,
  ToastBody,
  Spinner,
  useId,
  teamsDarkTheme,
  teamsHighContrastTheme,
  teamsLightTheme,
} from '@fluentui/react-components';
import type { PartialTheme } from '@fluentui/react-theme';
import { CopilotProvider } from '@fluentui-copilot/react-copilot';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { useConfigValidation } from '../../hooks/useConfigValidation';
import { useAgentHashRoute, slugify } from '../../hooks/useAgentHashRoute';
import { useAudienceFilter } from '../../hooks/useAudienceFilter';
import { useTeamsHost } from '../../hooks/useTeamsHost';
import { ConfigRequiredMessage } from '../config/ConfigRequiredMessage';
import { PersonalAppSettingsDrawer } from '../config/PersonalAppSettingsDrawer';
import { loadPersonalAppAgents, savePersonalAppAgents } from '../../config/personalAppConfig';
import { ChatLayout, IChatMessage } from './ChatLayout';
import type { IFoundryAgentsProps } from './IFoundryAgentsProps';
import type { IAgentConfig } from '../../models';
import { FoundryAgentService, isConversationGone, type IConversationSummary, type IMessageDetailsData } from '../../services/FoundryAgentService';
import { AgentStreamManager } from '../../services/AgentStreamManager';
import { DocumentService } from '../../services/DocumentService';
import {
  documentFromToolOutput,
  friendlyToolLabel,
  type IConsentRequest,
  type IDocumentRef,
  type IMcpServerState
} from '../../services/mcpTypes';
import { McpConnectionsButton, ConsentCard, openConsentPopup, serverDisplayName } from '../mcp/McpConnections';
import { DocumentPanel } from '../document/DocumentPanel';
import { Button as FluentButton, Tooltip as FluentTooltip } from '@fluentui/react-components';
import { DocumentText20Regular } from '@fluentui/react-icons';
import { createTranscriber } from '../../hooks/transcriber';
import {
  resolveAdditionalInformationTemplate,
  getDefaultDateTimeContext
} from '../../utils/additionalInformationResolver';
import strings from 'FoundryAgentsWebPartStrings';
import { tokens } from '@fluentui/react-components';
import styles from './FoundryAgents.module.scss';

/** Teams theme string from app.getContext() */
const TEAMS_THEME_DARK = 'dark';
const TEAMS_THEME_CONTRAST = 'contrast';

/**
 * Resolves Fluent theme for the current host. In Teams, the app runs in an iframe with no
 * SharePoint theme context, so we must pass an explicit theme for FluentProvider to work.
 * In SharePoint (full page / web part) we leave theme undefined so the host can inherit.
 */
function useResolvedFluentTheme(context: IFoundryAgentsProps['context']): PartialTheme | undefined {
  const [teamsTheme, setTeamsTheme] = React.useState<string>('default');
  const [isTeams, setIsTeams] = React.useState<boolean>(false);
  const mounted = React.useRef(true);

  const runTeamsThemeSetup = React.useCallback(
    (teamsJs: {
      app: {
        getContext(): Promise<{ app?: { theme?: string } }>;
        registerOnThemeChangeHandler?(handler: (theme: string) => void): void;
      };
    }) => {
      setIsTeams(true);
      teamsJs.app.getContext().then((ctx) => {
        if (mounted.current) {
          setTeamsTheme(ctx.app?.theme ?? 'default');
        }
      }).catch(() => {
        if (mounted.current) {
          setTeamsTheme('default');
        }
      });
      const register = teamsJs.app.registerOnThemeChangeHandler;
      if (typeof register === 'function') {
        register((theme: string) => {
          if (mounted.current) setTeamsTheme(theme);
        });
      }
    },
    []
  );

  React.useEffect(() => {
    mounted.current = true;
    const sdks = context.sdks as
      | {
          microsoftTeams?: {
            teamsJs?: {
              app?: {
                getContext(): Promise<{ app?: { theme?: string } }>;
                registerOnThemeChangeHandler?(handler: (theme: string) => void): void;
              };
            };
          };
        }
      | undefined;
    const teamsJs = sdks?.microsoftTeams?.teamsJs;
    if (teamsJs?.app) {
      runTeamsThemeSetup({ app: teamsJs.app });
      return () => {
        mounted.current = false;
      };
    }
    const sdksAsync = (context as { sdksAsync?: Promise<unknown> }).sdksAsync;
    if (sdksAsync) {
      sdksAsync.then((resolved) => {
        const asyncTeams = (resolved as { microsoftTeams?: { teamsJs?: { app?: unknown } } })?.microsoftTeams;
        if (asyncTeams?.teamsJs?.app && mounted.current) {
          runTeamsThemeSetup({ app: asyncTeams.teamsJs.app as Parameters<typeof runTeamsThemeSetup>[0]['app'] });
        }
      }).catch(() => undefined);
    }
    return () => {
      mounted.current = false;
    };
  }, [context, runTeamsThemeSetup]);

  // In Teams we must pass an explicit theme so FluentProvider injects styles into the iframe.
  if (!isTeams) {
    return undefined;
  }
  if (teamsTheme === TEAMS_THEME_DARK) {
    return teamsDarkTheme;
  }
  if (teamsTheme === TEAMS_THEME_CONTRAST) {
    return teamsHighContrastTheme;
  }
  return teamsLightTheme;
}

export default function FoundryAgents(props: IFoundryAgentsProps): React.ReactElement {
  const { agents, onOpenPropertyPane, context, minHeight = 'min(600px, 90vh)', openSettingsTrigger } = props;
  const { isTeamsTab, isPersonalApp, isTeams } = useTeamsHost(context);
  const fluentTheme = useResolvedFluentTheme(context);
  const showPersonalAppDialog = isPersonalApp || (isTeams && !isTeamsTab);

  const [personalAppAgents, setPersonalAppAgents] = React.useState<IAgentConfig[]>([]);
  const [openPersonalAppSettings, setOpenPersonalAppSettings] = React.useState(false);
  const [extensionLoadStatus, setExtensionLoadStatus] = React.useState<'idle' | 'loading' | 'success' | 'error'>('idle');
  const [extensionLoadError, setExtensionLoadError] = React.useState<unknown>(null);
  const [personalAppSaveError, setPersonalAppSaveError] = React.useState<unknown>(null);
  const personalAppToasterId = useId('foundry-personal-app-toaster');
  const { dispatchToast } = useToastController(personalAppToasterId);

  React.useEffect(() => {
    if (!isPersonalApp) return;
    setExtensionLoadStatus('loading');
    let cancelled = false;
    loadPersonalAppAgents(context)
      .then((list) => {
        if (!cancelled) {
          setPersonalAppAgents(list);
          setExtensionLoadStatus('success');
          setExtensionLoadError(null);
        }
      })
      .catch((err) => {
        if (!cancelled) {
          setPersonalAppAgents([]);
          setExtensionLoadStatus('error');
          setExtensionLoadError(err);
          const msg = err instanceof Error ? err.message : String(err);
          dispatchToast(
            <Toast appearance="inverted">
              <ToastTitle>{strings.PersonalAppLoadErrorTitle}</ToastTitle>
              <ToastBody>{msg}</ToastBody>
            </Toast>
          );
        }
      });
    return () => {
      cancelled = true;
    };
  }, [isPersonalApp, context, dispatchToast]);

  const agentsToUse = isPersonalApp ? personalAppAgents : agents;
  const { isValid } = useConfigValidation(agentsToUse, strings);

  // ── Audience filtering ────────────────────────────────────────────────────
  // `visibleAgents` is the audience-filtered subset. Membership checks are cached
  // so subsequent renders (e.g. agent switch) are instant.
  const { visibleAgents, status: audienceStatus } = useAudienceFilter(agentsToUse, context);

  const handleOpenPropertyPane = React.useCallback(() => {
    if (showPersonalAppDialog) {
      setOpenPersonalAppSettings(true);
    } else {
      onOpenPropertyPane();
    }
  }, [showPersonalAppDialog, onOpenPropertyPane]);

  React.useEffect(() => {
    if (openSettingsTrigger !== undefined && openSettingsTrigger !== null && openSettingsTrigger !== 0 && showPersonalAppDialog) {
      setOpenPersonalAppSettings(true);
    }
  }, [openSettingsTrigger, showPersonalAppDialog]);

  // ── Hash-based URL routing ────────────────────────────────────────────────
  const hashRoute = useAgentHashRoute();

  const [selectedAgent, setSelectedAgent] = React.useState<IAgentConfig | undefined>(undefined);
  const [messages, setMessages] = React.useState<IChatMessage[]>([]);
  const [isLoading, setIsLoading] = React.useState(false);
  const [isStreaming, setIsStreaming] = React.useState(false);
  const [conversationId, setConversationId] = React.useState<string | undefined>(undefined);
  const [conversations, setConversations] = React.useState<IConversationSummary[]>([]);
  const [conversationsNextCursor, setConversationsNextCursor] = React.useState<string | undefined>(undefined);
  const [isLoadingConversations, setIsLoadingConversations] = React.useState(false);
  const [isLoadingMoreConversations, setIsLoadingMoreConversations] = React.useState(false);
  const [hasMoreMessages, setHasMoreMessages] = React.useState(false);
  const [isLoadingMoreMessages, setIsLoadingMoreMessages] = React.useState(false);
  const [messagesCursor, setMessagesCursor] = React.useState<string | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [userDisplayName, setUserDisplayName] = React.useState<string>('');
  const [userMail, setUserMail] = React.useState<string>('');
  const [userPrincipalName, setUserPrincipalName] = React.useState<string>('');
  const [conversationSearchQuery, setConversationSearchQuery] = React.useState<string>('');
  const [toolRunning, setToolRunning] = React.useState<string | null>(null);
  const [streamState, setStreamState] = React.useState<'idle' | 'thinking' | 'executing' | 'done'>('idle');
  const abortControllerRef = React.useRef<AbortController | null>(null);
  const lastSentRef = React.useRef<{ text: string; files?: Array<{ name: string; fileId: string }> } | null>(null);
  type ConversationCacheEntry = { messages: IChatMessage[]; hasMore: boolean; cursor?: string };
  const conversationCacheRef = React.useRef<Map<string, ConversationCacheEntry>>(new Map());
  /** Cache conversations list per agent so history is instant when opening panel or switching agent */
  const conversationsListCacheRef = React.useRef<Map<string, { items: IConversationSummary[]; nextCursor?: string }>>(new Map());
  /**
   * Track the last agentId/conversationId we pushed to the URL ourselves so the
   * back/forward effect can distinguish external URL changes from our own.
   */
  const lastPushedAgentIdRef = React.useRef<string | undefined>(undefined);
  const lastPushedConversationIdRef = React.useRef<string | undefined>(undefined);
  /** Last completed response id per conversation — required for next message to avoid 400 (no tool output). */
  const lastResponseIdByConvRef = React.useRef<Map<string, string>>(new Map());

  // ── MCP connections (per user, per agent) ──────────────────────────────────
  const [mcpServers, setMcpServers] = React.useState<IMcpServerState[]>([]);
  const [checkingMcp, setCheckingMcp] = React.useState(false);
  const [consentBusy, setConsentBusy] = React.useState(false);
  const agentApiNameRef = React.useRef<Map<string, string>>(new Map());

  // ── Word documents produced by the agent (right panel) ─────────────────────
  const [documents, setDocuments] = React.useState<IDocumentRef[]>([]);
  const [activeDocId, setActiveDocId] = React.useState<string | undefined>(undefined);
  const [docPanelOpen, setDocPanelOpen] = React.useState(false);

  React.useEffect(() => {
    const factory = (context as unknown as { msGraphClientFactory?: { getClient: (version: '3') => Promise<MSGraphClientV3> } }).msGraphClientFactory;
    const pageUser = context?.pageContext?.user as { displayName?: string; email?: string; loginName?: string } | undefined;
    if (!factory) {
      if (pageUser?.displayName) setUserDisplayName(pageUser.displayName);
      if (pageUser?.email) setUserMail(pageUser.email);
      if (pageUser?.loginName) setUserPrincipalName(pageUser.loginName);
      return;
    }
    factory
      .getClient('3')
      .then((client: MSGraphClientV3) =>
        client.api('/me').select('givenName,displayName,mail,userPrincipalName').get()
      )
      .then((me: { givenName?: string; displayName?: string; mail?: string; userPrincipalName?: string }) => {
        const name = (me?.givenName ?? me?.displayName ?? '').trim();
        if (name) setUserDisplayName(name);
        if (me?.mail) setUserMail(me.mail);
        if (me?.userPrincipalName) setUserPrincipalName(me.userPrincipalName);
      })
      .catch(() => {
        if (pageUser?.displayName) setUserDisplayName(pageUser.displayName);
        if (pageUser?.email) setUserMail(pageUser.email);
        if (pageUser?.loginName) setUserPrincipalName(pageUser.loginName);
      });
  }, [context]);

  const dictationLang = React.useMemo(() => {
    const culture = (context?.pageContext as { cultureInfo?: { currentUICultureName?: string } } | undefined)?.cultureInfo?.currentUICultureName;
    return culture || 'en-US';
  }, [context?.pageContext]);

  const currentSiteUrl = React.useMemo(
    () => (context?.pageContext as { web?: { absoluteUrl?: string } } | undefined)?.web?.absoluteUrl ?? '',
    [context?.pageContext]
  );

  /** SharePoint tenant origin for spo_* function calls (current site only, no custom tenant). */
  const tenantHost = React.useMemo(() => {
    if (!currentSiteUrl) return '';
    try {
      return new URL(currentSiteUrl).origin;
    } catch {
      return '';
    }
  }, [currentSiteUrl]);

  const getTokenFn = React.useMemo(() => {
    return async (resource: string): Promise<string> => {
      const provider = await context.aadTokenProviderFactory.getTokenProvider();
      const token = await provider.getToken(resource);
      return token ?? '';
    };
  }, [context.aadTokenProviderFactory]);

  const serviceByEndpoint = React.useRef<Record<string, FoundryAgentService>>({});
  const getService = React.useCallback(
    (agent: IAgentConfig): FoundryAgentService => {
      const endpoint = agent.endpointUrl.replace(/\/$/, '');
      if (!serviceByEndpoint.current[endpoint]) {
        serviceByEndpoint.current[endpoint] = new FoundryAgentService(endpoint, getTokenFn);
      }
      return serviceByEndpoint.current[endpoint];
    },
    [getTokenFn]
  );

  const docService = React.useMemo(() => new DocumentService(getTokenFn), [getTokenFn]);
  // Server transcription when the agent has a transcription endpoint (same Entra token as the agents API);
  // otherwise the voice recorder uses the browser speech recognition.
  const transcriptionUrl = selectedAgent?.transcriptionUrl?.trim();
  const transcriptionPrompt = selectedAgent?.transcriptionPrompt;
  const transcriber = React.useMemo(
    () => (transcriptionUrl ? createTranscriber(getTokenFn, transcriptionUrl, transcriptionPrompt) : undefined),
    [transcriptionUrl, transcriptionPrompt, getTokenFn]
  );

  const resolveAgentName = React.useCallback(async (agent: IAgentConfig): Promise<string> => {
    const cached = agentApiNameRef.current.get(agent.agentId);
    if (cached) return cached;
    const name = await getService(agent).getResolvedAgentName(agent.agentId);
    agentApiNameRef.current.set(agent.agentId, name);
    return name;
  }, [getService]);

  const mcpCacheKey = (agentId: string): string => `foundryAgents.mcp.${agentId}`;
  const saveMcpCache = React.useCallback((agentId: string, list: IMcpServerState[]) => {
    try { localStorage.setItem(mcpCacheKey(agentId), JSON.stringify(list.map(({ consentLink, ...rest }) => rest))); } catch { /* storage unavailable */ }
  }, []);

  const updateMcpServer = React.useCallback((label: string, patch: Partial<IMcpServerState>) => {
    setMcpServers((prev) => {
      if (!prev.some((x) => x.label === label)) return prev;
      const next = prev.map((x) => (x.label === label ? { ...x, ...patch, checkedAt: Date.now() } : x));
      if (selectedAgent) saveMcpCache(selectedAgent.agentId, next);
      return next;
    });
  }, [selectedAgent, saveMcpCache]);

  const checkMcp = React.useCallback(async () => {
    if (!selectedAgent || !mcpServers.length) return;
    setCheckingMcp(true);
    try {
      const name = await resolveAgentName(selectedAgent);
      const states = await getService(selectedAgent).probeMcp(name, mcpServers);
      setMcpServers(states);
      saveMcpCache(selectedAgent.agentId, states);
    } catch {
      /* keep previous state */
    } finally {
      setCheckingMcp(false);
    }
  }, [selectedAgent, mcpServers, resolveAgentName, getService, saveMcpCache]);

  // Load the agent's MCP servers (from its Foundry definition) and the cached status, then check once.
  React.useEffect(() => {
    setMcpServers([]);
    if (!selectedAgent) return;
    let cancelled = false;
    (async () => {
      try {
        const name = await resolveAgentName(selectedAgent);
        const service = getService(selectedAgent);
        const servers = await service.getAgentMcpServers(name);
        if (cancelled || !servers.length) return;
        let cached: IMcpServerState[] = [];
        try { cached = JSON.parse(localStorage.getItem(mcpCacheKey(selectedAgent.agentId)) ?? '[]'); } catch { cached = []; }
        const initial = servers.map((sv) => {
          const c = cached.find((x) => x.label === sv.label);
          return c ? { ...sv, status: c.status, toolCount: c.toolCount, checkedAt: c.checkedAt } : { ...sv, status: 'unknown' as const };
        });
        setMcpServers(initial);
        const stale = initial.some((x) => x.status !== 'connected' || !x.checkedAt || Date.now() - x.checkedAt > 30 * 60_000);
        if (stale) {
          const states = await service.probeMcp(name, servers);
          if (!cancelled) {
            setMcpServers(states);
            saveMcpCache(selectedAgent.agentId, states);
          }
        }
      } catch {
        /* agent without MCP or not reachable */
      }
    })().catch(() => undefined);
    return () => { cancelled = true; };
  }, [selectedAgent?.agentId, resolveAgentName, getService, saveMcpCache]);

  const upsertDocument = React.useCallback((doc: IDocumentRef, open: boolean) => {
    setDocuments((prev) => {
      const others = prev.filter((d) => d.itemId !== doc.itemId);
      const old = prev.find((d) => d.itemId === doc.itemId);
      return [...others, { ...old, ...doc, updatedAt: doc.updatedAt ?? old?.updatedAt }];
    });
    if (open) {
      setActiveDocId(doc.itemId);
      setDocPanelOpen(true);
    }
  }, []);

  React.useEffect(() => {
    // Wait for audience checks to complete before selecting an agent
    if (audienceStatus === 'loading') return;

    if (isValid && visibleAgents && visibleAgents.length > 0) {
      if (!selectedAgent) {
        // On first load prefer the agent indicated by the URL hash, fallback to first.
        const urlAgent = hashRoute.agentSlug
          ? visibleAgents.find((a) => slugify(a.displayName) === hashRoute.agentSlug)
          : undefined;
        setSelectedAgent(urlAgent ?? visibleAgents[0]);
      } else {
        // Keep selectedAgent in sync with the latest data.
        const updated = visibleAgents.find((a) => a.agentId === selectedAgent.agentId);
        if (updated && updated !== selectedAgent) {
          setSelectedAgent(updated);
        } else if (!updated) {
          // Previously selected agent removed or now restricted – fallback.
          const urlAgent = hashRoute.agentSlug
            ? visibleAgents.find((a) => slugify(a.displayName) === hashRoute.agentSlug)
            : undefined;
          setSelectedAgent(urlAgent ?? visibleAgents[0]);
        }
      }
    } else if (!isValid || !visibleAgents || visibleAgents.length === 0) {
      setSelectedAgent(undefined);
    }
    // hashRoute.agentId intentionally excluded: this effect only uses it on first
    // load / after agent list changes. Browser back/forward is handled separately.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isValid, visibleAgents, selectedAgent, audienceStatus]);

  const cancelStream = React.useCallback(() => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }
  }, []);

  // Latest messages, read when a conversation lost by Foundry has to be recreated.
  const messagesRef = React.useRef<IChatMessage[]>([]);
  messagesRef.current = messages;

  const sendMessage = React.useCallback(
    async (text: string, uploadedFiles?: Array<{ name: string; fileId: string }>) => {
      if (!selectedAgent || !text.trim()) return;
      const userContent = text.trim();
      lastSentRef.current = { text: userContent, files: uploadedFiles };
      setError(undefined);

      abortControllerRef.current?.abort();
      abortControllerRef.current = new AbortController();
      const signal = abortControllerRef.current.signal;

      setMessages((prev) => [
        ...prev,
        { role: 'user' as const, content: userContent, uploadedFiles: uploadedFiles ?? [] },
        { role: 'assistant' as const, content: '', steps: [], startedAt: Date.now() }
      ]);
      setIsStreaming(true);
      setIsLoading(true);
      setToolRunning(null);
      setStreamState('idle');

      try {
        const service = getService(selectedAgent);

        // AgentStreamManager: uses service for conversation, stream, and tool continuation (no 400 on second message)
        const agentNameForApi = await resolveAgentName(selectedAgent);
        const resolvedAdditional = resolveAdditionalInformationTemplate(
          selectedAgent.additionalInformationTemplate,
          {
            userDisplayName,
            userMail,
            userPrincipalName,
            ...getDefaultDateTimeContext()
          }
        );
        const additionalContextPrefix =
          resolvedAdditional ? `[Context: ${resolvedAdditional}]` : undefined;
        // Context block given to the agent: user, date and the conversation id (stored in the documents it creates,
        // so the Word add-in can resume the same conversation).
        const buildContextPrefix = (convId: string): string => {
          let tz = '';
          try { tz = Intl.DateTimeFormat().resolvedOptions().timeZone ?? ''; } catch { tz = ''; }
          const now = new Date().toLocaleString('fr-CH', { dateStyle: 'full', timeStyle: 'short' });
          const parts = [
            `utilisateur: ${userDisplayName || ''}${userMail ? ` (${userMail})` : ''}`,
            `date: ${now}${tz ? ` (${tz})` : ''}`,
            `conversation_id: ${convId}`,
            `agent: ${agentNameForApi}`
          ];
          if (currentSiteUrl) parts.push(`site: ${currentSiteUrl}`);
          if (resolvedAdditional) parts.push(resolvedAdditional);
          return `[Contexte] ${parts.join(' ; ')}`;
        };
        const attachToLastAssistant = (patch: (m: IChatMessage) => IChatMessage): void => {
          setMessages((prev) => {
            const next = [...prev];
            const last = next[next.length - 1];
            if (last?.role === 'assistant') next[next.length - 1] = patch(last);
            return next;
          });
        };

        const manager = new AgentStreamManager({
          service,
          agentNameForApi,
          tenantHost,
          signal,
          allowFunctionCalls: selectedAgent.allowFunctionCalls === true,
          blockedHttpMethods: selectedAgent.blockedHttpMethods,
          siteContext: currentSiteUrl ? { siteUrl: currentSiteUrl } : undefined,
          ownerUpn: userPrincipalName || undefined,
          additionalContextPrefix,
          buildContextPrefix,
          callbacks: {
            onStateChange: (state, payload) => {
              setStreamState(state);
              if (state === 'executing' && payload?.functionNames?.length) {
                setToolRunning(friendlyToolLabel(payload.functionNames[0] ?? ''));
              } else if (state !== 'executing') {
                setToolRunning(null);
              }
            },
            onConversation: (convId) => {
              setConversationId(convId);
            },
            onConsentRequest: (consent) => {
              attachToLastAssistant((m) => ({ ...m, consent }));
              updateMcpServer(consent.serverLabel, { status: 'consent_required', consentLink: consent.link, toolCount: 0 });
            },
            onMcpListTools: (label, count, err) => {
              if (count > 0) updateMcpServer(label, { status: 'connected', toolCount: count, consentLink: undefined, error: undefined });
              else if (err) updateMcpServer(label, { status: 'error', error: err });
            },
            onStep: (step) => {
              attachToLastAssistant((m) => {
                const steps = [...(m.steps ?? [])];
                const idx = steps.findIndex((x) => x.id === step.id);
                if (idx >= 0) steps[idx] = step; else steps.push(step);
                return { ...m, steps };
              });
            },
            onMcpCall: (call, phase) => {
              if (phase !== 'done') return;
              const doc = documentFromToolOutput(call.name, call.output);
              if (!doc) return;
              upsertDocument(doc, call.name !== 'doc_read');
              attachToLastAssistant((m) => ({
                ...m,
                documents: [...(m.documents ?? []).filter((d) => d.itemId !== doc.itemId), doc]
              }));
            },
            onChunk: (chunk) => {
              setMessages((prev) => {
                const next = [...prev];
                const last = next[next.length - 1];
                if (last?.role === 'assistant') {
                  next[next.length - 1] = { ...last, content: last.content + chunk };
                }
                return next;
              });
            }
          }
        });

        const lastResponseId = conversationId ? lastResponseIdByConvRef.current.get(conversationId) : undefined;
        // Earlier exchanges shown on screen (before this turn), used if Foundry lost the conversation.
        const earlier = messagesRef.current
          .filter((m) => (m.role === 'user' || m.role === 'assistant') && m.content?.trim())
          .slice(-20)
          .map((m) => ({ role: m.role as 'user' | 'assistant', text: m.content.slice(0, 6000) }));
        let result: Awaited<ReturnType<typeof manager.run>>;
        let recreatedFrom: string | undefined;
        try {
          result = await manager.run(conversationId, userContent, lastResponseId);
        } catch (runErr) {
          // Foundry sometimes loses a conversation (404 / conversation_not_found, see learn.microsoft.com Q&A 5859393):
          // recreate it with the exchanges still on screen and send the message again, instead of a raw error.
          if (!conversationId || !isConversationGone(runErr) || signal.aborted) throw runErr;
          recreatedFrom = conversationId;
          lastResponseIdByConvRef.current.delete(conversationId);
          conversationCacheRef.current.delete(conversationId);
          attachToLastAssistant((m) => ({ ...m, content: '', steps: [] }));
          result = await manager.run(undefined, userContent, undefined, earlier);
        }
        const { conversationId: newConvId, outputText, lastResponseId: nextLastResponseId, consent } = result;
        if (recreatedFrom) {
          const goneId = recreatedFrom;
          setConversations((prev) => prev.filter((c) => c.id !== goneId));
          const cachedList = conversationsListCacheRef.current.get(selectedAgent.agentId);
          if (cachedList) {
            conversationsListCacheRef.current.set(selectedAgent.agentId, { ...cachedList, items: cachedList.items.filter((c) => c.id !== goneId) });
          }
        }

        const wasNewConversation = !conversationId || !!recreatedFrom;
        setConversationId(newConvId);
        lastPushedAgentIdRef.current = slugify(selectedAgent.displayName);
        lastPushedConversationIdRef.current = newConvId;
        if (nextLastResponseId) {
          lastResponseIdByConvRef.current.set(newConvId, nextLastResponseId);
        }
        hashRoute.navigateTo(selectedAgent.displayName, newConvId);

        if (wasNewConversation) {
          const title = userContent.trim().slice(0, 50) || strings.NewConversationTitle;
          const newEntry: IConversationSummary = { id: newConvId, title, updatedAt: new Date() };
          setConversations((prev) =>
            prev.some((t) => t.id === newConvId) ? prev : [newEntry, ...prev]
          );
          const cacheKey = selectedAgent.agentId;
          const cached = conversationsListCacheRef.current.get(cacheKey);
          if (cached && !cached.items.some((t) => t.id === newConvId)) {
            conversationsListCacheRef.current.set(cacheKey, {
              ...cached,
              items: [newEntry, ...cached.items]
            });
          }
        }

        setMessages((prev) => {
          const next = [...prev];
          const last = next[next.length - 1];
          const finalText = outputText || (consent ? strings.ConsentSignInToContinue.replace('{0}', serverDisplayName(consent.serverLabel)) : outputText);
          if (last?.role === 'assistant') {
            next[next.length - 1] = { ...last, content: finalText, endedAt: Date.now() };
          }
          return next;
        });

        service.getConversationItemsPaged(newConvId, { limit: 5 })
          .then(({ messages: recent }) => {
            const latestAssistant = [...recent].reverse().find(m => m.role === 'assistant');
            if (latestAssistant?.id) {
              setMessages(prev => {
                const next = [...prev];
                const last = next[next.length - 1];
                if (last?.role === 'assistant' && !last.id) {
                  next[next.length - 1] = { ...last, id: latestAssistant.id };
                }
                return next;
              });
            }
          })
          .catch(() => {/* non-critical */});

        // Update conversation metadata with last_item_id for notification badge
        service.getConversationItemsPaged(newConvId, { limit: 1 })
          .then(({ messages: latest }) => {
            const lastId = latest[0]?.id;
            if (lastId) {
              service.updateConversationMetadata(newConvId, { last_item_id: lastId }).then(() => {
                setConversations((p) => p.map((c) => (c.id === newConvId ? { ...c, lastItemId: lastId } : c)));
                const cacheKey = selectedAgent.agentId;
                const cached = conversationsListCacheRef.current.get(cacheKey);
                if (cached) {
                  conversationsListCacheRef.current.set(cacheKey, {
                    ...cached,
                    items: cached.items.map((c) => (c.id === newConvId ? { ...c, lastItemId: lastId } : c))
                  });
                }
              }).catch(() => {});
            }
          })
          .catch(() => {});
      } catch (err) {
        if ((err as Error)?.name === 'AbortError') {
          return;
        }
        // Strip the failed assistant placeholder so the error component takes its place
        setMessages((prev) => {
          const next = [...prev];
          if (next[next.length - 1]?.role === 'assistant') next.pop();
          return next;
        });
        setError(err instanceof Error ? err.message : String(err));
      } finally {
        setIsStreaming(false);
        setIsLoading(false);
        setToolRunning(null);
        setStreamState('idle');
        // eslint-disable-next-line require-atomic-updates
        abortControllerRef.current = null;
      }
    },
    [selectedAgent, getService, conversationId, hashRoute.navigateTo, tenantHost, currentSiteUrl, userDisplayName, userMail, userPrincipalName, resolveAgentName, updateMcpServer, upsertDocument]
  );

  /** In-chat sign-in: open Foundry's consent page, then resume the request automatically. */
  const handleConsentSignIn = React.useCallback(async (consent: IConsentRequest) => {
    setConsentBusy(true);
    try {
      await openConsentPopup(consent.link);
      setMessages((prev) => prev.map((m) => (m.consent?.link === consent.link ? { ...m, consent: { ...m.consent, done: true } } : m)));
      updateMcpServer(consent.serverLabel, { status: 'connected', consentLink: undefined });
    } finally {
      setConsentBusy(false);
    }
    await sendMessage(strings.ConsentResumeMessage.replace('{0}', serverDisplayName(consent.serverLabel)));
  }, [sendMessage, updateMcpServer]);

  /** Sign-in from the connections popover (outside a turn): then re-check the connections. */
  const handlePopoverSignIn = React.useCallback(async (server: IMcpServerState) => {
    if (!server.consentLink) return;
    await openConsentPopup(server.consentLink);
    await checkMcp();
  }, [checkMcp]);

  const handleRetry = React.useCallback(() => {
    if (!lastSentRef.current) return;
    const { text, files } = lastSentRef.current;
    // Remove the dangling user message so sendMessage can re-add it cleanly
    setMessages((prev) => {
      if (prev.length > 0 && prev[prev.length - 1].role === 'user') {
        return prev.slice(0, -1);
      }
      return prev;
    });
    setError(undefined);
    sendMessage(text, files).catch(() => undefined);
  }, [sendMessage]);

  const handleStarterPromptClick = React.useCallback(
    (prompt: string) => {
      if (selectedAgent) sendMessage(prompt).catch(() => {});
    },
    [selectedAgent, sendMessage]
  );

  const loadConversations = React.useCallback(
    async (options?: { skipFetchIfCached?: boolean }) => {
      if (!selectedAgent) return;
      const cacheKey = selectedAgent.agentId;
      const cached = conversationsListCacheRef.current.get(cacheKey);
      if (cached) {
        setConversations(cached.items);
        setConversationsNextCursor(cached.nextCursor);
        if (options?.skipFetchIfCached) return;
      }
      setIsLoadingConversations(true);
      try {
        const service = getService(selectedAgent);
        const { items, nextCursor } = await service.listConversations(selectedAgent.agentId, {
          limit: 100,
          ownerUpn: userPrincipalName || undefined,
          onlyWithTitleInMeta: selectedAgent.showOnlyWebPartConversations === true
        });
        setConversations(items);
        setConversationsNextCursor(nextCursor);
        conversationsListCacheRef.current.set(cacheKey, { items, nextCursor });
      } catch {
        if (!cached) {
          setConversations([]);
          setConversationsNextCursor(undefined);
        }
      } finally {
        setIsLoadingConversations(false);
      }
    },
    [selectedAgent, getService, userPrincipalName]
  );

  const loadMoreConversations = React.useCallback(async () => {
    if (!selectedAgent || !conversationsNextCursor || isLoadingMoreConversations) return;
    setIsLoadingMoreConversations(true);
    try {
      const service = getService(selectedAgent);
      const { items, nextCursor } = await service.listConversations(selectedAgent.agentId, {
        limit: 100,
          ownerUpn: userPrincipalName || undefined,
        after: conversationsNextCursor,
        onlyWithTitleInMeta: selectedAgent.showOnlyWebPartConversations === true
      });
      setConversations((prev) => [...prev, ...items]);
      setConversationsNextCursor(nextCursor);
    } catch {
      setConversationsNextCursor(undefined);
    } finally {
      setIsLoadingMoreConversations(false);
    }
  }, [selectedAgent, getService, conversationsNextCursor, isLoadingMoreConversations, userPrincipalName]);

  /** Preload conversation history on init and on agent switch (async, no wait for panel). */
  React.useEffect(() => {
    if (!selectedAgent) return;
    loadConversations().catch(() => undefined);
  }, [selectedAgent?.agentId, loadConversations]);

  // ── Back / forward navigation ─────────────────────────────────────────────
  // React to URL hash changes that were NOT triggered by our own handlers
  // (e.g. the user pressed the browser back / forward button).

  React.useEffect(() => {
    const { agentSlug: urlAgentSlug, conversationId: urlConversationId } = hashRoute;
    const sameAgent = urlAgentSlug === lastPushedAgentIdRef.current;
    const sameConversation = urlConversationId === lastPushedConversationIdRef.current;

    // Nothing relevant changed – skip.
    if (sameAgent && sameConversation) return;

    lastPushedAgentIdRef.current = urlAgentSlug;
    lastPushedConversationIdRef.current = urlConversationId;

    if (!isValid || !visibleAgents || !visibleAgents.length) return;

    const currentAgentSlug = selectedAgent ? slugify(selectedAgent.displayName) : undefined;

    // --- Agent changed via URL (browser back/forward) ---
    if (!sameAgent && urlAgentSlug && urlAgentSlug !== currentAgentSlug) {
      const urlAgent = visibleAgents.find((a) => slugify(a.displayName) === urlAgentSlug);
      if (urlAgent) {
        if (conversationId && messages.length > 0) {
          conversationCacheRef.current.set(conversationId, { messages, hasMore: hasMoreMessages, cursor: messagesCursor });
        }
        setConversationId(undefined);
        setMessages([]);
        setHasMoreMessages(false);
        setMessagesCursor(undefined);
        setSelectedAgent(urlAgent);
        // Conversation for the new agent will be handled in the next render cycle
        // once selectedAgent state propagates.
        return;
      }
    }

    // --- Conversation changed via URL (same agent, browser back/forward) ---
    if (!sameConversation && urlAgentSlug === currentAgentSlug) {
      if (!urlConversationId) {
        // Back to "no conversation" state.
        if (conversationId) {
          if (messages.length > 0) conversationCacheRef.current.set(conversationId, { messages, hasMore: hasMoreMessages, cursor: messagesCursor });
          setConversationId(undefined);
          setMessages([]);
          setHasMoreMessages(false);
          setMessagesCursor(undefined);
        }
      } else if (urlConversationId !== conversationId && selectedAgent) {
        // Navigate to a specific conversation via back/forward.
        setError(undefined);
        const cached = conversationCacheRef.current.get(urlConversationId);
        if (cached) {
          setConversationId(urlConversationId);
          setMessages(cached.messages);
          setHasMoreMessages(cached.hasMore);
          setMessagesCursor(cached.cursor);
        } else {
          setConversationId(urlConversationId);
          setHasMoreMessages(false);
          setMessagesCursor(undefined);
          setIsLoading(true);
          const service = getService(selectedAgent);
          service
            .getConversationItemsPaged(urlConversationId, { limit: 20 })
            .then(({ messages: items, hasMore, nextCursor }) => {
              setMessages(items);
              setHasMoreMessages(hasMore);
              setMessagesCursor(nextCursor);
              conversationCacheRef.current.set(urlConversationId, { messages: items, hasMore, cursor: nextCursor });
            })
            .catch((e: unknown) => {
              setMessages([]);
              setHasMoreMessages(false);
              if (!isConversationGone(e)) return;
              // Link to a conversation Foundry no longer has (bookmark, restored tab): clear it and say so.
              setConversationId(undefined);
              setConversations((prev) => prev.filter((c) => c.id !== urlConversationId));
              lastPushedConversationIdRef.current = undefined;
              hashRoute.navigateTo(selectedAgent.displayName);
              lastSentRef.current = null;
              setError('Cette conversation n\'existe plus dans Azure AI Foundry. Vous pouvez en commencer une nouvelle.');
            })
            .finally(() => setIsLoading(false));
        }
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hashRoute.agentSlug, hashRoute.conversationId]);

  // ── Audience enforcement for direct-link navigation ──────────────────────
  // After audience checks complete, verify that the agent in the URL (if any) is
  // actually accessible. If not, redirect to the first visible agent so the user
  // cannot reach a restricted agent by pasting a direct link.
  const didEnforceAudienceRef = React.useRef(false);
  React.useEffect(() => {
    if (audienceStatus !== 'ready') return;
    if (didEnforceAudienceRef.current) return;
    if (!hashRoute.agentSlug || !isValid) return;

    // Find the agent the URL targets (ignoring audience)
    const targetAgent = agents.find((a) => slugify(a.displayName) === hashRoute.agentSlug);
    if (!targetAgent) return;

    // Check if it is in the visible (audience-allowed) list
    const isAllowed = visibleAgents.some((a) => a.agentId === targetAgent.agentId);
    if (!isAllowed) {
      // Access denied – silently redirect to the first visible agent (or clear URL)
      didEnforceAudienceRef.current = true;
      const fallback = visibleAgents[0];
      if (fallback) {
        lastPushedAgentIdRef.current = slugify(fallback.displayName);
        lastPushedConversationIdRef.current = undefined;
        hashRoute.navigateTo(fallback.displayName);
        setSelectedAgent(fallback);
      } else {
        window.location.hash = '';
        setSelectedAgent(undefined);
      }
    } else {
      didEnforceAudienceRef.current = true;
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [audienceStatus, isValid]);

  // ── Initial conversation restore from URL ────────────────────────────────
  // When the page loads with a URL that contains a conversationId, load the
  // conversation once the selected agent becomes available. The
  // `didLoadInitialConversationRef` prevents subsequent agent switches from
  // re-triggering this logic.
  const didLoadInitialConversationRef = React.useRef(false);
  React.useEffect(() => {
    if (didLoadInitialConversationRef.current) return;
    if (!hashRoute.conversationId || !selectedAgent || slugify(selectedAgent.displayName) !== hashRoute.agentSlug) return;
    didLoadInitialConversationRef.current = true;
    const convId = hashRoute.conversationId;
    // Stamp refs so the back/forward effect won't double-load.
    lastPushedAgentIdRef.current = slugify(selectedAgent.displayName);
    lastPushedConversationIdRef.current = convId;
    const cached = conversationCacheRef.current.get(convId);
    if (cached) {
      setConversationId(convId);
      setMessages(cached.messages);
      setHasMoreMessages(cached.hasMore);
      setMessagesCursor(cached.cursor);
    } else {
      setConversationId(convId);
      setHasMoreMessages(false);
      setMessagesCursor(undefined);
      setIsLoading(true);
      const service = getService(selectedAgent);
      service
        .getConversationItemsPaged(convId, { limit: 20 })
        .then(({ messages: items, hasMore, nextCursor }) => {
          setMessages(items);
          setHasMoreMessages(hasMore);
          setMessagesCursor(nextCursor);
          conversationCacheRef.current.set(convId, { messages: items, hasMore, cursor: nextCursor });
        })
        .catch((e: unknown) => {
          setMessages([]);
          setHasMoreMessages(false);
          if (!isConversationGone(e)) return;
          // Link to a conversation Foundry no longer has (bookmark, restored tab): clear it and say so.
          setConversationId(undefined);
          setConversations((prev) => prev.filter((c) => c.id !== convId));
          lastPushedConversationIdRef.current = undefined;
          hashRoute.navigateTo(selectedAgent.displayName);
          lastSentRef.current = null;
          setError('Cette conversation n\'existe plus dans Azure AI Foundry. Vous pouvez en commencer une nouvelle.');
        })
        .finally(() => setIsLoading(false));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedAgent?.agentId]);

  const resetDocuments = React.useCallback(() => {
    setDocuments([]);
    setActiveDocId(undefined);
    setDocPanelOpen(false);
  }, []);

  // Documents of the open conversation (doc_* tool calls stored by Foundry).
  const loadedDocsForConvRef = React.useRef<string | undefined>(undefined);
  React.useEffect(() => {
    if (!conversationId || !selectedAgent || isStreaming) return;
    if (loadedDocsForConvRef.current === conversationId) return;
    loadedDocsForConvRef.current = conversationId;
    getService(selectedAgent).getConversationDocuments(conversationId)
      .then((docs) => {
        if (!docs.length) return;
        setDocuments((prev) => {
          const ids = new Set(prev.map((d) => d.itemId));
          return [...docs.filter((d) => !ids.has(d.itemId)), ...prev];
        });
      })
      .catch(() => { /* optional */ });
  }, [conversationId, selectedAgent, isStreaming, getService]);

  const handleNewConversation = React.useCallback(() => {
    setError(undefined);
    resetDocuments();
    loadedDocsForConvRef.current = undefined;
    if (conversationId && messages.length > 0) {
      conversationCacheRef.current.set(conversationId, { messages, hasMore: hasMoreMessages, cursor: messagesCursor });
    }
    setConversationId(undefined);
    setMessages([]);
    setHasMoreMessages(false);
    setMessagesCursor(undefined);
    // Clear the conversation segment from the URL (keep agent).
    if (selectedAgent) {
      lastPushedAgentIdRef.current = slugify(selectedAgent.displayName);
      lastPushedConversationIdRef.current = undefined;
      hashRoute.navigateTo(selectedAgent.displayName);
    }
  }, [conversationId, messages, hasMoreMessages, messagesCursor, selectedAgent, hashRoute.navigateTo, resetDocuments]);

  /** On agent switch: clear chat like "New conversation", then set the new agent. No-op if same agent. */
  const handleSelectAgent = React.useCallback((agent: IAgentConfig) => {
    setError(undefined);
    if (selectedAgent?.agentId === agent.agentId) return;
    resetDocuments();
    loadedDocsForConvRef.current = undefined;
    if (conversationId && messages.length > 0) {
      conversationCacheRef.current.set(conversationId, { messages, hasMore: hasMoreMessages, cursor: messagesCursor });
    }
    setConversationId(undefined);
    setMessages([]);
    setHasMoreMessages(false);
    setMessagesCursor(undefined);
    setSelectedAgent(agent);
    // Reflect the agent switch in the URL (no active conversation yet).
    lastPushedAgentIdRef.current = slugify(agent.displayName);
    lastPushedConversationIdRef.current = undefined;
    hashRoute.navigateTo(agent.displayName);
  }, [selectedAgent?.agentId, conversationId, messages, hasMoreMessages, messagesCursor, hashRoute.navigateTo, resetDocuments]);

  const handleDeleteConversation = React.useCallback(
    async (convId: string) => {
      if (!selectedAgent) return;
      const service = getService(selectedAgent);
      await service.deleteConversation(convId);
      setConversations((prev) => prev.filter((t) => t.id !== convId));
      conversationCacheRef.current.delete(convId);
      const cacheKey = selectedAgent.agentId;
      const cached = conversationsListCacheRef.current.get(cacheKey);
      if (cached) {
        conversationsListCacheRef.current.set(cacheKey, {
          ...cached,
          items: cached.items.filter((t) => t.id !== convId)
        });
      }
      if (conversationId === convId) {
        setConversationId(undefined);
        setMessages([]);
        lastPushedAgentIdRef.current = slugify(selectedAgent.displayName);
        lastPushedConversationIdRef.current = undefined;
        hashRoute.navigateTo(selectedAgent.displayName);
      }
    },
    [selectedAgent, getService, conversationId, hashRoute.navigateTo]
  );

  const updateLastUsedSiteIfNeeded = React.useCallback(
    (convId: string, agent: IAgentConfig | undefined, conversationList: IConversationSummary[]) => {
      if (!currentSiteUrl || !agent) return;
      const summary = conversationList.find((c) => c.id === convId);
      if (!summary) return;
      const previousSite = summary.lastUsedSiteUrl ?? summary.creationSiteUrl;
      if (previousSite === currentSiteUrl) return;
      const existingMeta: Record<string, unknown> = {
        ...(summary.rawMeta ?? {}),
        title: summary.title,
        ...(summary.creationSiteUrl && { creation_site_url: summary.creationSiteUrl }),
        ...(summary.lastUsedSiteUrl && { last_used_site_url: summary.lastUsedSiteUrl }),
        ...(summary.lastItemId && { last_item_id: summary.lastItemId }),
        ...(summary.lastSeenItemId && { last_seen_item_id: summary.lastSeenItemId })
      };
      getService(agent)
        .updateConversationMetadata(convId, {
          last_used_site_url: currentSiteUrl,
          last_used_at: Math.floor(Date.now() / 1000)
        }, existingMeta)
        .then(() => {
          setConversations((p) => p.map((c) => (c.id === convId ? { ...c, lastUsedSiteUrl: currentSiteUrl } : c)));
          const cacheKey = agent.agentId;
          const cached = conversationsListCacheRef.current.get(cacheKey);
          if (cached) {
            conversationsListCacheRef.current.set(cacheKey, {
              ...cached,
              items: cached.items.map((c) => (c.id === convId ? { ...c, lastUsedSiteUrl: currentSiteUrl } : c))
            });
          }
        })
        .catch(() => { /* non-fatal */ });
    },
    [currentSiteUrl, getService]
  );

  const handleSelectConversation = React.useCallback(
    async (convId: string) => {
      // A notice about the previous conversation must not stay on the one being opened.
      setError(undefined);
      if (convId !== conversationId) {
        resetDocuments();
        loadedDocsForConvRef.current = undefined;
      }
      const cached = conversationCacheRef.current.get(convId);
      if (cached) {
        setConversationId(convId);
        setMessages(cached.messages);
        setHasMoreMessages(cached.hasMore);
        setMessagesCursor(cached.cursor);
        if (selectedAgent) {
          lastPushedAgentIdRef.current = slugify(selectedAgent.displayName);
          lastPushedConversationIdRef.current = convId;
          hashRoute.navigateTo(selectedAgent.displayName, convId);
        }
        updateLastUsedSiteIfNeeded(convId, selectedAgent, conversations);
        const lastSeenId = cached.messages[cached.messages.length - 1]?.id;
        if (lastSeenId && selectedAgent) {
          const summary = conversations.find((c) => c.id === convId);
          const existingMeta: Record<string, unknown> | undefined = summary
            ? {
                ...(summary.rawMeta ?? {}),
                title: summary.title,
                ...(summary.creationSiteUrl && { creation_site_url: summary.creationSiteUrl }),
                ...(summary.lastUsedSiteUrl && { last_used_site_url: summary.lastUsedSiteUrl }),
                ...(summary.lastItemId && { last_item_id: summary.lastItemId }),
                ...(summary.lastSeenItemId && { last_seen_item_id: summary.lastSeenItemId })
              }
            : undefined;
          getService(selectedAgent)
            .updateConversationMetadata(convId, {
              last_seen_item_id: lastSeenId,
              last_seen_at: Math.floor(Date.now() / 1000)
            }, existingMeta)
            .then(() => {
              setConversations((p) => p.map((c) => (c.id === convId ? { ...c, lastSeenItemId: lastSeenId } : c)));
              const cacheKey = selectedAgent.agentId;
              const cachedList = conversationsListCacheRef.current.get(cacheKey);
              if (cachedList) {
                conversationsListCacheRef.current.set(cacheKey, {
                  ...cachedList,
                  items: cachedList.items.map((c) => (c.id === convId ? { ...c, lastSeenItemId: lastSeenId } : c))
                });
              }
            })
            .catch(() => {});
        }
        return;
      }
      if (!selectedAgent) return;
      setConversationId(convId);
      setHasMoreMessages(false);
      setMessagesCursor(undefined);
      // Reflect conversation selection in the URL immediately (before messages load).
      lastPushedAgentIdRef.current = slugify(selectedAgent.displayName);
      lastPushedConversationIdRef.current = convId;
      hashRoute.navigateTo(selectedAgent.displayName, convId);
      updateLastUsedSiteIfNeeded(convId, selectedAgent, conversations);
      setIsLoading(true);
      try {
        const service = getService(selectedAgent);
        let page: Awaited<ReturnType<typeof service.getConversationItemsPaged>>;
        try {
          page = await service.getConversationItemsPaged(convId, { limit: 20 });
        } catch (loadErr) {
          if (!isConversationGone(loadErr)) throw loadErr;
          // Lost by Foundry: drop it from the history and say so instead of showing an empty chat.
          setConversations((prev) => prev.filter((c) => c.id !== convId));
          const cachedList = conversationsListCacheRef.current.get(selectedAgent.agentId);
          if (cachedList) {
            conversationsListCacheRef.current.set(selectedAgent.agentId, { ...cachedList, items: cachedList.items.filter((c) => c.id !== convId) });
          }
          setConversationId(undefined);
          setMessages([]);
          lastPushedConversationIdRef.current = undefined;
          hashRoute.navigateTo(selectedAgent.displayName);
          lastSentRef.current = null; // nothing to retry: the error only reports the missing conversation
          setError('Cette conversation n\'existe plus dans Azure AI Foundry : elle a été retirée de l\'historique. Vous pouvez en commencer une nouvelle.');
          return;
        }
        const { messages: items, hasMore, nextCursor } = page;
        setMessages(items);
        setHasMoreMessages(hasMore);
        setMessagesCursor(nextCursor);
        conversationCacheRef.current.set(convId, { messages: items, hasMore, cursor: nextCursor });
        const lastSeenId = items[items.length - 1]?.id;
        if (lastSeenId) {
          const summary = conversations.find((c) => c.id === convId);
          const existingMeta: Record<string, unknown> | undefined = summary
            ? {
                ...(summary.rawMeta ?? {}),
                title: summary.title,
                ...(summary.creationSiteUrl && { creation_site_url: summary.creationSiteUrl }),
                ...(summary.lastUsedSiteUrl && { last_used_site_url: summary.lastUsedSiteUrl }),
                ...(summary.lastItemId && { last_item_id: summary.lastItemId }),
                ...(summary.lastSeenItemId && { last_seen_item_id: summary.lastSeenItemId })
              }
            : undefined;
          service
            .updateConversationMetadata(convId, {
              last_seen_item_id: lastSeenId,
              last_seen_at: Math.floor(Date.now() / 1000)
            }, existingMeta)
            .then(() => {
              setConversations((p) => p.map((c) => (c.id === convId ? { ...c, lastSeenItemId: lastSeenId } : c)));
              const cacheKey = selectedAgent.agentId;
              const cached = conversationsListCacheRef.current.get(cacheKey);
              if (cached) {
                conversationsListCacheRef.current.set(cacheKey, {
                  ...cached,
                  items: cached.items.map((c) => (c.id === convId ? { ...c, lastSeenItemId: lastSeenId } : c))
                });
              }
            })
            .catch(() => {});
        }
      } catch {
        setMessages([]);
        setHasMoreMessages(false);
      } finally {
        setIsLoading(false);
      }
    },
    [selectedAgent, getService, hashRoute.navigateTo, conversations, updateLastUsedSiteIfNeeded, conversationId, resetDocuments]
  );

  const loadMoreMessages = React.useCallback(async () => {
    if (!selectedAgent || !conversationId || !hasMoreMessages || isLoadingMoreMessages) return;
    setIsLoadingMoreMessages(true);
    try {
      const service = getService(selectedAgent);
      const { messages: older, hasMore, nextCursor } = await service.getConversationItemsPaged(
        conversationId,
        { limit: 100, after: messagesCursor }
      );
      setMessages(prev => {
        const merged = [...older, ...prev];
        conversationCacheRef.current.set(conversationId, { messages: merged, hasMore, cursor: nextCursor });
        return merged;
      });
      setHasMoreMessages(hasMore);
      setMessagesCursor(nextCursor);
    } catch {
      // keep current messages
    } finally {
      setIsLoadingMoreMessages(false);
    }
  }, [selectedAgent, conversationId, hasMoreMessages, isLoadingMoreMessages, messagesCursor, getService]);

  /**
   * Fetch full details for an assistant message (model, reasoning, tool calls, usage).
   * Called by the "⋯" button on history messages that carry a server-side item id.
   */
  const handleGetMessageDetails = React.useCallback(
    async (messageId: string): Promise<IMessageDetailsData | null> => {
      if (!selectedAgent || !conversationId) return null;
      const service = getService(selectedAgent);
      return service.getMessageDetails(conversationId, messageId);
    },
    [selectedAgent, conversationId, getService]
  );

  const conversationsDisplay = React.useMemo((): IConversationSummary[] => {
    const list = [...conversations];
    if (conversationId && messages.length > 0) {
      const firstUser = messages.find((m) => m.role === 'user');
      const title = (firstUser?.content ?? '').trim().slice(0, 50) || strings.NewConversationTitle;
      if (!list.some((t) => t.id === conversationId)) {
        list.unshift({ id: conversationId, title, updatedAt: new Date() });
      }
    }
    return list;
  }, [conversations, conversationId, messages]);

  const conversationsFiltered = React.useMemo((): IConversationSummary[] => {
    const q = conversationSearchQuery.trim().toLowerCase();
    if (!q) return conversationsDisplay;
    return conversationsDisplay.filter((t) => t.title.toLowerCase().includes(q));
  }, [conversationsDisplay, conversationSearchQuery]);

  // Brand color for Waves so they use Teams theme in Teams (DOM theme vars may not be ready yet).
  const wavesBrandColor = fluentTheme
    ? ((fluentTheme as Record<string, string>).colorBrandBackground ?? (fluentTheme as Record<string, string>).colorBrandStroke1)
    : undefined;

  const content = isPersonalApp && extensionLoadStatus === 'loading' ? (
    <div className={styles.content} style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: tokens.spacingVerticalM }}>
      <Spinner size="medium" />
      <span>{strings.LoadingPersonalAppSettings}</span>
    </div>
  ) : !isValid ? (
    <div className={styles.content}>
      <ConfigRequiredMessage
        onOpenPropertyPane={handleOpenPropertyPane}
        description={isPersonalApp ? strings.PersonalAppMustConfigure : strings.AdminMustConfigure}
        teamsTabConfigureHint={isTeamsTab ? strings.TeamsTabConfigureHint : undefined}
        wavesBrandColor={wavesBrandColor}
      />
    </div>
  ) : (
    <div className={styles.content}>
      <ChatLayout
          agents={visibleAgents}
          wavesBrandColor={wavesBrandColor}
          selectedAgent={selectedAgent}
          messages={messages}
          onSelectAgent={handleSelectAgent}
          onSend={sendMessage}
          onUploadFile={selectedAgent ? (file) => getService(selectedAgent).uploadFile(file) : undefined}
          onStarterPromptClick={handleStarterPromptClick}
          userLastName={userDisplayName}
          isLoading={isLoading}
          isStreaming={isStreaming}
          onCancelStream={cancelStream}
          error={error}
          onRetry={handleRetry}
          conversations={conversationsFiltered}
          onSelectConversation={handleSelectConversation}
          onDeleteConversation={handleDeleteConversation}
          onNewConversation={handleNewConversation}
          conversationSearchQuery={conversationSearchQuery}
          onConversationSearchQueryChange={setConversationSearchQuery}
          onOpenHistoryDrawer={() => loadConversations({ skipFetchIfCached: true })}
          isLoadingConversations={isLoadingConversations}
          hasMoreConversations={!!conversationsNextCursor}
          loadMoreConversations={loadMoreConversations}
          isLoadingMoreConversations={isLoadingMoreConversations}
          selectedConversationId={conversationId}
          allowAttachments={selectedAgent?.allowAttachments === true}
          hasMoreMessages={hasMoreMessages}
          loadMoreMessages={loadMoreMessages}
          isLoadingMoreMessages={isLoadingMoreMessages}
          onGetMessageDetails={handleGetMessageDetails}
          toolName={toolRunning ?? undefined}
          streamState={streamState}
          showPersonalAppSettingsButton={isPersonalApp}
          onOpenPersonalAppSettings={isPersonalApp ? () => setOpenPersonalAppSettings(true) : undefined}
          dictationLang={dictationLang}
          transcribe={transcriber}
          headerExtras={
            <>
              {documents.length > 0 && !docPanelOpen && (
                <FluentTooltip content={strings.ShowDocument} relationship="label">
                  <FluentButton appearance="subtle" shape="circular" icon={<DocumentText20Regular />} onClick={() => setDocPanelOpen(true)}>
                    {documents.length > 1 ? strings.DocumentsCount.replace('{0}', String(documents.length)) : strings.DocumentSingle}
                  </FluentButton>
                </FluentTooltip>
              )}
              <McpConnectionsButton
                servers={mcpServers}
                checking={checkingMcp}
                onCheck={() => { checkMcp().catch(() => undefined); }}
                onSignIn={(sv) => { handlePopoverSignIn(sv).catch(() => undefined); }}
              />
            </>
          }
          renderMessageExtras={(msg) => (
            <>
              {msg.consent && (
                <ConsentCard consent={msg.consent} busy={consentBusy} onSignIn={(c) => { handleConsentSignIn(c).catch(() => undefined); }} />
              )}
              {msg.documents?.map((d) => (
                <div key={d.itemId} className={styles.docChip}>
                  <FluentButton
                    appearance="outline"
                    icon={<DocumentText20Regular />}
                    onClick={() => { setActiveDocId(d.itemId); setDocPanelOpen(true); }}
                  >
                    {d.name}
                  </FluentButton>
                </div>
              ))}
            </>
          )}
          rightPanel={docPanelOpen && documents.length > 0 ? (
            <DocumentPanel
              documents={documents}
              activeItemId={activeDocId}
              service={docService}
              onSelect={setActiveDocId}
              onClose={() => setDocPanelOpen(false)}
            />
          ) : undefined}
        />
    </div>
  );

  // Workaround for SharePoint FluentProvider ID collision (sp-dev-docs #9847): isolate our
  // FluentProvider with a unique ID prefix so generated IDs don't clash with SP's FluentUI.
  const targetDoc = typeof document !== 'undefined' ? document : undefined;
  // Single portal mount inside FluentProvider so Drawer, Tooltip, MenuPopover inherit theme.
  const [portalMount, setPortalMount] = React.useState<HTMLElement | null>(null);

  return (
    <IdPrefixProvider value="foundry-agents-chat">
      <FluentProvider
        targetDocument={targetDoc}
        theme={fluentTheme}
        style={{ display: 'flex', flex: 1, height: '100%' }}
        applyStylesToPortals
      >
        <div ref={setPortalMount} style={{ position: 'relative', display: 'flex', flex: 1, height: '100%' }} />
        <PortalMountNodeProvider value={portalMount ?? undefined}>
          <div className={styles.foundryAgents} style={{ minHeight }}>
            <CopilotProvider style={{ display: 'flex', flex: 1 }}>
              {content}
            </CopilotProvider>
          </div>
          {isPersonalApp && <Toaster toasterId={personalAppToasterId} position="bottom-end" />}
          {showPersonalAppDialog && (
            <PersonalAppSettingsDrawer
              open={openPersonalAppSettings}
              onOpenChange={setOpenPersonalAppSettings}
              agents={personalAppAgents}
                onSave={async (newAgents) => {
                try {
                  await savePersonalAppAgents(context, JSON.stringify(newAgents));
                  setPersonalAppAgents(newAgents);
                  setPersonalAppSaveError(null);
                } catch (err) {
                  setPersonalAppSaveError(err);
                }
              }}
              extensionError={extensionLoadError}
              saveError={personalAppSaveError}
              portalMountNode={isPersonalApp ? undefined : (portalMount ?? undefined)}
              context={context}
              strings={{
                PersonalAppSettingsLabel: strings.PersonalAppSettingsLabel,
                DetailsClose: strings.DetailsClose,
              }}
            />
          )}
        </PortalMountNodeProvider>
      </FluentProvider>
    </IdPrefixProvider>
  );
}
