/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import {
  MessageBar,
  MessageBarBody,
  Button,
  SearchBox,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Tooltip,
  Spinner,
  Skeleton,
  SkeletonItem,
  Toaster,
  useToastController,
  Toast,
  ToastTitle,
  useId,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions
} from '@fluentui/react-components';
import {
  ChevronDown24Regular,
  History24Regular,
  Add24Regular,
  Dismiss24Regular,
  Attach24Regular,
  Delete20Regular,
  Document20Regular,
  DocumentPdf20Regular,
  DocumentCsv20Regular,
  DocumentText20Regular,
  DocumentJavascript20Regular,
  DocumentJava20Regular,
  DocumentCss20Regular,
  DocumentSass20Regular,
  DocumentYml20Regular,
  Image20Regular,
  Video20Regular,
  Archive20Regular,
  FolderZip20Regular,
  Code20Regular,
  CommentMultipleRegular,
  ErrorCircle20Filled,
  Settings20Regular,
  Mic24Regular,
  ArrowDown20Regular
} from '@fluentui/react-icons';
import {
  ChatInput as CopilotChatInput,
  ImperativeControlPlugin,
  type ImperativeControlPluginRef
} from '@fluentui-copilot/react-copilot';
import { Attachment, AttachmentList } from '@fluentui-copilot/react-attachments';
import strings from 'FoundryAgentsWebPartStrings';
import type { IAgentConfig } from '../../models';
import type { IChatMessage } from '../../types/chat';
import type { IMessageDetailsData } from '../../services/FoundryAgentService';
import { AgentIcon } from './shared/AgentIcon';
import { StarterMessages } from './shared/StarterMessages';
import { Waves } from './shared/Waves';
import { AssistantMessage } from './messages/AssistantMessage';
import { UserMessage } from './messages/UserMessage';
import { AgentActivity } from './messages/AgentActivity';
import styles from './ChatLayout.module.scss';
import { useVoiceRecorder } from '../../hooks/useVoiceRecorder';
import { VoiceRecorderBar } from './shared/VoiceRecorderBar';

export type { IChatMessage };

/** Summary of a conversation for the history panel */
export interface IConversationSummary {
  id: string;
  title: string;
  updatedAt?: Date;
  /** For notification badge: last message id in conversation */
  lastItemId?: string;
  /** Last message id the user has seen */
  lastSeenItemId?: string;
}

/** A file that has been picked by the user and is being (or has been) uploaded. */
export interface IAttachedFile {
  /** Stable local key used for React rendering and dismissal. */
  localId: string;
  file: File;
  /** Set once the upload succeeds. */
  fileId?: string;
  uploading: boolean;
  error: boolean;
  /** Human-readable error message to show in the tooltip when upload fails. */
  errorMessage?: string;
}

export interface IChatLayoutProps {
  agents: IAgentConfig[];
  selectedAgent: IAgentConfig | undefined;
  messages: IChatMessage[];
  /** Called on submit. uploadedFiles contains only files whose upload succeeded. */
  onSend: (text: string, uploadedFiles?: Array<{ name: string; fileId: string }>) => void;
  /** Upload a file immediately when the user attaches it. Returns the file_id. */
  onUploadFile?: (file: File) => Promise<string>;
  onStarterPromptClick?: (prompt: string) => void;
  onSelectAgent: (agent: IAgentConfig) => void;
  /** Current user's last name for the welcome message (from loc WelcomeMessage) */
  userLastName?: string;
  isLoading?: boolean;
  isStreaming?: boolean;
  onCancelStream?: () => void;
  error?: string | undefined;
  onRetry?: () => void;
  /** Conversations for the history panel */
  conversations?: IConversationSummary[];
  /** Called when user selects a conversation from history */
  onSelectConversation?: (conversationId: string) => void;
  /** Called when user clicks "New conversation" */
  onNewConversation?: () => void;
  /** Called when the history rail is opened (e.g. to load first page of conversations) */
  onOpenHistoryDrawer?: () => void;
  /** Whether conversations are currently loading from the service */
  isLoadingConversations?: boolean;
  /** Whether there are more conversations to load (for infinite scroll) */
  hasMoreConversations?: boolean;
  /** Load next page of conversations (called when sentinel is visible) */
  loadMoreConversations?: () => void;
  /** Whether the next page of conversations is loading */
  isLoadingMoreConversations?: boolean;
  /** Current conversation id (for highlighting in rail) */
  selectedConversationId?: string;
  /** Search query for filtering conversations in the rail (client-side filter by title) */
  conversationSearchQuery?: string;
  /** Called when the user changes the conversation search query */
  onConversationSearchQueryChange?: (value: string) => void;
  /** Whether the attachment button is shown in the chat input */
  allowAttachments?: boolean;
  /** Called when user clicks the delete button on a conversation. Must return a Promise so the toast reflects the actual outcome. */
  onDeleteConversation?: (conversationId: string) => Promise<void>;
  /** Whether there are older messages to load (triggers top sentinel) */
  hasMoreMessages?: boolean;
  /** Load the previous page of messages (called when top sentinel is visible) */
  loadMoreMessages?: () => void;
  /** Whether an older-message page is being fetched */
  isLoadingMoreMessages?: boolean;
  /**
   * Fetch full details (model, reasoning, tool calls, token usage) for an
   * assistant message by its server-side item ID.
   * When provided, a "⋯" button is shown on assistant messages that have an ID.
   */
  // eslint-disable-next-line @rushstack/no-new-null
  onGetMessageDetails?: (messageId: string) => Promise<IMessageDetailsData | null>;
  /** When set (e.g. in Teams), waves use this brand color for theme consistency */
  wavesBrandColor?: string;
  /** Current tool name when streamState is 'executing' (for loader label) */
  toolName?: string;
  /** Stream state for UI: idle | thinking | executing | done */
  streamState?: 'idle' | 'thinking' | 'executing' | 'done';
  /** When true (e.g. Teams Personal App), show a settings icon button in the header next to New conversation */
  showPersonalAppSettingsButton?: boolean;
  /** Called when the personal app settings button is clicked */
  onOpenPersonalAppSettings?: () => void;
  /** Extra controls in the header (e.g. MCP connections button). */
  headerExtras?: React.ReactNode;
  /** Content rendered under an assistant message (sign-in card, document chips). */
  renderMessageExtras?: (message: IChatMessage, index: number) => React.ReactNode;
  /** Right-hand panel (document preview). */
  rightPanel?: React.ReactNode;
  /** Voice dictation language (BCP-47); dictation button hidden when undefined or unsupported. */
  dictationLang?: string;
  /** Server transcription of the dictated audio (gpt-4o-transcribe); the browser live text is used without it. */
  transcribe?: (audio: Blob, lang: string) => Promise<string>;
}


/** Returns the most specific Fluent icon component for a given File's MIME type and extension. */
function getFileTypeIcon(file: File): React.ReactElement {
  const mime = (file.type ?? '').toLowerCase();
  const ext = (file.name.split('.').pop() ?? '').toLowerCase();

  if (mime === 'application/pdf' || ext === 'pdf') return <DocumentPdf20Regular />;
  if (mime === 'text/csv' || ext === 'csv') return <DocumentCsv20Regular />;
  if (mime.startsWith('image/')) return <Image20Regular />;
  if (mime.startsWith('video/')) return <Video20Regular />;
  if (
    mime === 'application/javascript' ||
    mime === 'application/typescript' ||
    ext === 'js' || ext === 'mjs' || ext === 'cjs' ||
    ext === 'jsx' || ext === 'ts' || ext === 'tsx'
  ) return <DocumentJavascript20Regular />;
  if (ext === 'java') return <DocumentJava20Regular />;
  if (mime === 'text/css' || ext === 'css') return <DocumentCss20Regular />;
  if (ext === 'scss' || ext === 'sass' || ext === 'less') return <DocumentSass20Regular />;
  if (ext === 'yaml' || ext === 'yml') return <DocumentYml20Regular />;
  if (
    ext === 'zip' || ext === 'gz' || ext === 'tar' || ext === 'bz2' ||
    ext === '7z' || ext === 'rar' || ext === 'tgz'
  ) return <FolderZip20Regular />;
  if (
    mime.includes('zip') || mime.includes('archive') || mime.includes('compressed')
  ) return <Archive20Regular />;
  if (
    ext === 'json' || ext === 'jsonl' || ext === 'xml' ||
    ext === 'html' || ext === 'htm' || ext === 'sh' ||
    ext === 'py' || ext === 'rb' || ext === 'go' ||
    ext === 'rs' || ext === 'swift' || ext === 'cpp' ||
    ext === 'c' || ext === 'h' || ext === 'cs' ||
    mime.startsWith('text/x-') || mime === 'application/json' || mime === 'application/xml'
  ) return <Code20Regular />;
  if (mime.startsWith('text/') || ext === 'txt' || ext === 'md' || ext === 'log') {
    return <DocumentText20Regular />;
  }
  return <Document20Regular />;
}

export interface IAgentHeaderProps {
  agents: IAgentConfig[];
  selectedAgent: IAgentConfig | undefined;
  onSelectAgent: (agent: IAgentConfig) => void;
}

export const AgentHeader: React.FunctionComponent<IAgentHeaderProps> = ({
  agents,
  selectedAgent,
  onSelectAgent
}) => {
  if (!selectedAgent) return null;
  if (agents.length <= 1) {
    return (
      <div className={styles.agentHeaderSingle}>
        <AgentIcon agent={selectedAgent} size={25} />
        <span>{selectedAgent.displayName}</span>
      </div>
    );
  }
  return (
    <Menu>
      <MenuTrigger>
        <button
          type="button"
          className={styles.agentSelectorTrigger}
          aria-label={strings.SelectedAgentAriaLabel.replace('{0}', selectedAgent.displayName)}
        >
          <AgentIcon agent={selectedAgent} size={25} />
          <span>{selectedAgent.displayName}</span>
          <ChevronDown24Regular />
        </button>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {agents.map((a) => (
            <MenuItem
              key={a.agentId}
              disabled={selectedAgent?.displayName === a.displayName}
              onClick={() => onSelectAgent(a)}
              icon={<AgentIcon agent={a} size={20} />}
            >
              {a.displayName}
            </MenuItem>
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

export const ChatLayout: React.FunctionComponent<IChatLayoutProps> = (props) => {
  const {
    agents,
    selectedAgent,
    messages,
    onSelectAgent,
    onSend,
    onUploadFile,
    onStarterPromptClick,
    toolName,
    streamState = 'idle',
    userLastName,
    isStreaming = false,
    onCancelStream,
    error,
    onRetry,
    conversations = [],
    onSelectConversation,
    onNewConversation,
    onOpenHistoryDrawer,
    isLoadingConversations = false,
    hasMoreConversations = false,
    loadMoreConversations,
    isLoadingMoreConversations = false,
    selectedConversationId,
    conversationSearchQuery = '',
    onConversationSearchQueryChange,
    allowAttachments = false,
    onDeleteConversation,
    hasMoreMessages = false,
    loadMoreMessages,
    isLoadingMoreMessages = false,
    onGetMessageDetails,
    wavesBrandColor,
    showPersonalAppSettingsButton = false,
    onOpenPersonalAppSettings,
    headerExtras,
    renderMessageExtras,
    rightPanel,
    dictationLang,
    transcribe,
  } = props;
  const toasterId = useId('foundry-toaster');
  const { dispatchToast } = useToastController(toasterId);
  const controlRef = React.useRef<ImperativeControlPluginRef>(null);
  const messagesEndRef = React.useRef<HTMLDivElement>(null);
  const messagesContainerRef = React.useRef<HTMLDivElement>(null);
  const scrollerRef = React.useRef<HTMLDivElement>(null);
  const contentRef = React.useRef<HTMLDivElement>(null);
  const inputWrapperRef = React.useRef<HTMLDivElement>(null);
  const msgTopSentinelRef = React.useRef<HTMLDivElement>(null);
  const sentinelRef = React.useRef<HTMLDivElement>(null);
  const railScrollRef = React.useRef<HTMLDivElement>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const isPrependingRef = React.useRef(false);
  const savedScrollRef = React.useRef<{ scrollTop: number; scrollHeight: number } | null>(null);
  const [railOpen, setRailOpen] = React.useState(false);
  const [attachedFiles, setAttachedFiles] = React.useState<IAttachedFile[]>([]);
  const [loadingConversationId, setLoadingConversationId] = React.useState<string | null>(null);
  const [conversationToDelete, setConversationToDelete] = React.useState<string | null>(null);
  // Voice input: server transcription first (browser recognition only as a silent backup), waveform + timer while recording.
  const voice = useVoiceRecorder(transcribe, dictationLang ?? 'fr-FR');
  const voiceActive = voice.state !== 'idle';
  const confirmVoice = React.useCallback(async () => {
    const text = await voice.stop();
    if (!text) return;
    const ctl = controlRef.current;
    const current = ctl?.getInputText?.() ?? '';
    if (current.trim()) ctl?.appendText?.(` ${text}`);
    else ctl?.setInputText?.(text);
  }, [voice]);

  const openRail = React.useCallback(() => {
    if (!railOpen) {
      onOpenHistoryDrawer?.();
      setRailOpen(true);
    } else {
      setRailOpen(false);
    }
  }, [onOpenHistoryDrawer, railOpen]);

  // Mirror exact message-top-sentinel IO pattern for conversations rail.
  React.useEffect(() => {
    if (!sentinelRef.current || !railOpen || !hasMoreConversations || isLoadingMoreConversations || !loadMoreConversations) return;
    const sentinel = sentinelRef.current;
    const root = railScrollRef.current ?? undefined;
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries[0]?.isIntersecting) loadMoreConversations();
      },
      { root, rootMargin: '80px', threshold: 0 }
    );
    observer.observe(sentinel);
    return () => observer.disconnect();
  }, [railOpen, hasMoreConversations, isLoadingMoreConversations, loadMoreConversations]);

  // ── Scrolling, Claude / ChatGPT style ────────────────────────────────────
  // The list follows the stream only while the user is at the bottom. Scrolling up (wheel, touch,
  // keys, scrollbar) detaches it; a "jump to bottom" button then appears. Sending a message,
  // opening a conversation or reaching the bottom again re-attaches it.
  const stickRef = React.useRef(true);
  const lastTopRef = React.useRef(0);
  const [showJump, setShowJump] = React.useState(false);
  const STICK_PX = 32;

  const distanceFromBottom = (el: HTMLElement): number => el.scrollHeight - el.scrollTop - el.clientHeight;

  const scrollToBottom = React.useCallback((smooth = false) => {
    const el = scrollerRef.current;
    if (!el) return;
    stickRef.current = true;
    setShowJump(false);
    el.scrollTo({ top: el.scrollHeight, behavior: smooth ? 'smooth' : ('instant' as ScrollBehavior) });
  }, []);

  const handleScroll = React.useCallback(() => {
    const el = scrollerRef.current;
    if (!el) return;
    const dist = distanceFromBottom(el);
    if (dist <= STICK_PX) stickRef.current = true;
    else if (el.scrollTop < lastTopRef.current - 1) stickRef.current = false; // moved up by the user
    lastTopRef.current = el.scrollTop;
    setShowJump(!stickRef.current && dist > 120);
  }, []);

  // Intent to read above: detach right away, before the next token pulls the view down.
  const detachOnUserUp = React.useCallback((e: React.WheelEvent | React.KeyboardEvent) => {
    if ('deltaY' in e ? e.deltaY < 0 : ['ArrowUp', 'PageUp', 'Home'].includes(e.key)) stickRef.current = false;
  }, []);
  const touchYRef = React.useRef(0);

  // Keep the bottom in view while attached, whatever grows: tokens, timeline steps, images, code blocks.
  React.useEffect(() => {
    const el = scrollerRef.current;
    const content = contentRef.current;
    if (!el || !content || typeof ResizeObserver === 'undefined') return;
    const ro = new ResizeObserver(() => {
      if (stickRef.current && !savedScrollRef.current) el.scrollTop = el.scrollHeight;
      else handleScroll();
    });
    ro.observe(content);
    return () => ro.disconnect();
  }, [handleScroll]);

  // Older messages prepended: keep the reading position (before paint, no jump).
  // New messages: follow only when attached; a message sent by the user always re-attaches.
  const prevLenRef = React.useRef(messages.length);
  React.useLayoutEffect(() => {
    const el = scrollerRef.current;
    const prevLen = prevLenRef.current;
    prevLenRef.current = messages.length;
    if (!el) return;
    if (savedScrollRef.current) {
      const heightAdded = el.scrollHeight - savedScrollRef.current.scrollHeight;
      if (heightAdded > 0) el.scrollTop = savedScrollRef.current.scrollTop + heightAdded;
      lastTopRef.current = el.scrollTop;
      savedScrollRef.current = null;
      isPrependingRef.current = false;
      return;
    }
    const last = messages[messages.length - 1];
    if (messages.length > prevLen && last?.role === 'user') stickRef.current = true;
    if (stickRef.current) el.scrollTop = el.scrollHeight;
  }, [messages]);

  // Another conversation (or agent) opened: start at the bottom, attached. Not when the id is only
  // being assigned to the conversation that is streaming right now (first message of a new chat).
  const streamingRef = React.useRef(isStreaming);
  streamingRef.current = isStreaming;
  React.useLayoutEffect(() => {
    if (!streamingRef.current) scrollToBottom(false);
  }, [selectedConversationId, selectedAgent?.agentId, scrollToBottom]);

  // Esc stops the answer being written (as in Claude / ChatGPT).
  React.useEffect(() => {
    if (!isStreaming || !onCancelStream) return;
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape' && !e.defaultPrevented && !document.querySelector('[role="dialog"]')) onCancelStream();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [isStreaming, onCancelStream]);

  // Focus the input when the answer ends and when switching agent / conversation.
  const focusInput = React.useCallback(() => {
    const target = inputWrapperRef.current?.querySelector<HTMLElement>('[contenteditable="true"], textarea');
    target?.focus({ preventScroll: true });
  }, []);
  const wasStreamingRef = React.useRef(isStreaming);
  React.useEffect(() => {
    if (wasStreamingRef.current && !isStreaming && !voiceActive) focusInput();
    wasStreamingRef.current = isStreaming;
  }, [isStreaming, voiceActive, focusInput]);
  React.useEffect(() => {
    if (selectedAgent) focusInput();
  }, [selectedConversationId, selectedAgent?.agentId, focusInput]);

  // IntersectionObserver on the top sentinel: load older messages when it enters view.
  React.useEffect(() => {
    if (!msgTopSentinelRef.current || !hasMoreMessages || isLoadingMoreMessages || !loadMoreMessages) return;
    const sentinel = msgTopSentinelRef.current;
    const root = scrollerRef.current ?? undefined;
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries[0]?.isIntersecting) {
          if (scrollerRef.current) {
            savedScrollRef.current = {
              scrollTop: scrollerRef.current.scrollTop,
              scrollHeight: scrollerRef.current.scrollHeight,
            };
          }
          isPrependingRef.current = true;
          loadMoreMessages();
        }
      },
      { root, rootMargin: '80px', threshold: 0 }
    );
    observer.observe(sentinel);
    return () => observer.disconnect();
  }, [hasMoreMessages, isLoadingMoreMessages, loadMoreMessages]);

  // Clear the per-conversation loading spinner once the conversation has been selected and data has loaded
  React.useEffect(() => {
    if (loadingConversationId && !props.isLoading && selectedConversationId === loadingConversationId) {
      setLoadingConversationId(null);
    }
  }, [props.isLoading, selectedConversationId, loadingConversationId]);

  const isUploading = attachedFiles.some((a) => a.uploading);

  const handleSubmit = React.useCallback(
    (_ev: unknown, data?: { value?: string }) => {
      const text = (data?.value ?? '').trim();
      if (!text || !selectedAgent) return;
      // Block send while any file is still being uploaded
      if (attachedFiles.some((a) => a.uploading)) return;
      const uploadedFiles = attachedFiles
        .filter((a) => !!a.fileId)
        .map((a) => ({ name: a.file.name, fileId: a.fileId! }));
      onSend(text, uploadedFiles.length > 0 ? uploadedFiles : undefined);
      setAttachedFiles([]);
      controlRef.current?.setInputText?.('');
    },
    [onSend, selectedAgent, attachedFiles]
  );

  const handleFileChange = React.useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const picked = Array.from(e.target.files ?? []);
      e.target.value = '';
      if (!picked.length) return;

      for (const file of picked) {
        const localId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
        // Add the chip immediately in "uploading" state
        setAttachedFiles((prev) => [
          ...prev,
          { localId, file, uploading: true, error: false }
        ]);
        // Fire upload right away
        (onUploadFile ? onUploadFile(file) : Promise.reject(new Error('no uploader')))
          .then((fileId) => {
            setAttachedFiles((prev) =>
              prev.map((a) =>
                a.localId === localId ? { ...a, fileId, uploading: false } : a
              )
            );
          })
          .catch((err: unknown) => {
            // Extract a human-readable message from the API error.
            // The SDK may throw with the raw JSON body or a parsed Error.message.
            let errorMessage = err instanceof Error ? err.message : String(err);
            try {
              // If the message is a JSON string containing { error: { message } }
              const parsed = JSON.parse(errorMessage) as { error?: { message?: string } };
              if (parsed?.error?.message) errorMessage = parsed.error.message;
            } catch { /* not JSON — use as-is */ }
            setAttachedFiles((prev) =>
              prev.map((a) =>
                a.localId === localId ? { ...a, uploading: false, error: true, errorMessage } : a
              )
            );
          });
      }
    },
    [onUploadFile]
  );

  const handleAttachmentDismiss = React.useCallback(
    (_e: React.MouseEvent, data: { id?: string }) => {
      setAttachedFiles((prev) => prev.filter((a) => a.localId !== data.id));
    },
    []
  );

  const handleDeleteConversation = React.useCallback(
    (conversationId: string) => {
      (onDeleteConversation?.(conversationId) ?? Promise.resolve())
        .then(() => {
          setConversationToDelete(null);
          dispatchToast(
            <Toast appearance="inverted">
              <ToastTitle>{strings.ConversationDeleted}</ToastTitle>
            </Toast>,
            { intent: 'success', timeout: 4000 }
          );
        })
        .catch(() => {
          dispatchToast(
            <Toast appearance="inverted">
              <ToastTitle>{strings.ConversationDeleteError}</ToastTitle>
            </Toast>,
            { intent: 'error', timeout: 6000 }
          );
        });
    },
    [onDeleteConversation, dispatchToast]
  );

  const handleConfirmDeleteConversation = React.useCallback(() => {
    if (conversationToDelete) {
      handleDeleteConversation(conversationToDelete);
    }
  }, [conversationToDelete, handleDeleteConversation]);

  return (
    <div className={styles.chatContainer}>
      <Dialog open={conversationToDelete !== null} onOpenChange={(_, data) => !data.open && setConversationToDelete(null)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{strings.DeleteConversationConfirmTitle}</DialogTitle>
            <DialogContent>{strings.DeleteConversationConfirmMessage}</DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setConversationToDelete(null)}>
                {strings.Cancel}
              </Button>
              <Button appearance="primary" onClick={handleConfirmDeleteConversation}>
                {strings.DeleteConversationConfirmButton}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      <div className={styles.layoutWithRail}>
        <aside className={styles.rail} data-open={railOpen} aria-label={strings.ConversationsTitle}>
          <div className={styles.railInner}>
            <div className={styles.railHeader}>
              <span className={styles.railTitle}>{strings.ConversationsTitle}</span>
              <Tooltip content={strings.ToggleConversations} relationship="label">
                <Button
                  appearance="subtle"
                  shape="circular"
                  size="small"
                  icon={<Dismiss24Regular />}
                  aria-label={strings.ToggleConversations}
                  onClick={() => setRailOpen(false)}
                />
              </Tooltip>
            </div>
            <div className={styles.railSearchBox}>
              <SearchBox
                placeholder={strings.SearchConversationsPlaceholder}
                value={conversationSearchQuery}
                onChange={(_e, data) => onConversationSearchQueryChange?.(data.value ?? '')}
                aria-label={strings.SearchConversationsPlaceholder}
              />
            </div>
            <div ref={railScrollRef} className={styles.railScroll} role="list">
              {isLoadingConversations ? (
                <div className={styles.railLoading}>
                  <Spinner size="tiny" />
                  <span>{strings.Loading}</span>
                </div>
              ) : conversations.length === 0 ? (
                <div className={styles.railEmpty}>
                  <CommentMultipleRegular className={styles.railEmptyIcon} />
                  <span>{strings.NoConversationsYet}</span>
                </div>
              ) : (
                <div className={styles.conversationsList}>
                  {conversations.map((t) => {
                    const isConversationLoading = loadingConversationId === t.id;
                    const hasUnread =
                      (t.lastItemId !== undefined && t.lastItemId !== null) &&
                      (t.lastSeenItemId === undefined || t.lastSeenItemId !== t.lastItemId);
                    return (
                      <div key={t.id} className={styles.conversationItem}>
                        <button
                          type="button"
                          className={styles.conversationCard}
                          data-selected={t.id === selectedConversationId}
                          disabled={isConversationLoading}
                          onClick={() => {
                            if (t.id !== selectedConversationId) {
                              setLoadingConversationId(t.id);
                            }
                            onSelectConversation?.(t.id);
                          }}
                          title={t.title}
                        >
                          <span className={styles.conversationCardTitle}>{t.title}</span>
                          {t.updatedAt && (
                            <span className={styles.conversationCardDate}>
                              {t.updatedAt.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })}
                            </span>
                          )}
                        </button>
                        {isConversationLoading ? (
                          <Spinner size="extra-small" className={styles.conversationLoadingSpinner} aria-label={strings.LoadingConversation} />
                        ) : (
                          <>
                            {hasUnread && (
                              <Tooltip content={strings.NotificationTooltip} relationship="label">
                                <span
                                  className={styles.conversationNotificationDot}
                                  style={{ backgroundColor: wavesBrandColor ?? 'var(--colorBrandBackground, #0078d4)' }}
                                  aria-hidden
                                />
                              </Tooltip>
                            )}
                            {onDeleteConversation ? (
                              <Tooltip content={strings.DeleteConversationConfirmTitle} relationship="label">
                                <Button
                                  appearance="subtle"
                                  shape="circular"
                                  size="small"
                                  icon={<Delete20Regular />}
                                  className={styles.conversationDeleteBtn}
                                  aria-label={strings.DeleteConversationConfirmTitle}
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setConversationToDelete(t.id);
                                  }}
                                />
                              </Tooltip>
                            ) : null}
                          </>
                        )}
                      </div>
                    );
                  })}
                  {hasMoreConversations && <div ref={sentinelRef} className={styles.sentinel} aria-hidden="true" />}
                  {isLoadingMoreConversations && (
                    <Skeleton className={styles.railShimmer} aria-hidden="true">
                      <div className={styles.railShimmerItem}>
                        <SkeletonItem shape="rectangle" size={16} className={styles.railShimmerTitle} />
                        <SkeletonItem shape="rectangle" size={12} className={styles.railShimmerDate} />
                      </div>
                    </Skeleton>
                  )}
                </div>
              )}
            </div>
          </div>
        </aside>
        <div className={styles.mainRow} data-panel={rightPanel ? 'true' : undefined}>
        <div className={styles.mainColumn}>
          {selectedAgent && (
            <header className={styles.header}>
              <div className={styles.headerLeft}>
                <Tooltip content={strings.ToggleConversations} relationship="label">
                  <Button
                    appearance="subtle"
                    shape="circular"
                    icon={<History24Regular />}
                    aria-label={strings.ToggleConversations}
                    onClick={openRail}
                  />
                </Tooltip>
                <AgentHeader
                agents={agents}
                selectedAgent={selectedAgent}
                onSelectAgent={onSelectAgent}
              />
              </div>
              <div className={styles.headerRight}>
              {headerExtras}
              {showPersonalAppSettingsButton && onOpenPersonalAppSettings && (
                <Tooltip content={strings.PersonalAppSettingsLabel} relationship="label">
                  <Button
                    appearance="subtle"
                    shape="circular"
                    icon={<Settings20Regular />}
                    aria-label={strings.PersonalAppSettingsLabel}
                    onClick={onOpenPersonalAppSettings}
                  />
                </Tooltip>
              )}
              {onNewConversation && (
                  <Button
                    appearance="subtle"
                    shape="circular"
                    icon={<Add24Regular />}
                    onClick={onNewConversation}
                  >
                    {strings.NewConversation}
                  </Button>
                )}
              </div>
            </header>
          )}
      <div ref={messagesContainerRef} className={styles.messagesContainer}>
        <div
          ref={scrollerRef}
          className={styles.messagesWrapper}
          data-conversation-loading={loadingConversationId ? 'true' : undefined}
          role="log"
          aria-live="polite"
          aria-busy={isStreaming || undefined}
          aria-label={strings.ChatMessagesAriaLabel}
          tabIndex={-1}
          onScroll={handleScroll}
          onWheel={detachOnUserUp}
          onKeyDown={detachOnUserUp}
          onTouchStart={(e) => { touchYRef.current = e.touches[0]?.clientY ?? 0; }}
          onTouchMove={(e) => { if ((e.touches[0]?.clientY ?? 0) > touchYRef.current + 4) stickRef.current = false; }}
        >
        <div ref={contentRef} className={styles.messagesColumn}>
          {(hasMoreMessages || isLoadingMoreMessages) && (
            <div ref={msgTopSentinelRef} className={styles.msgTopSentinel} aria-hidden="true">
              {isLoadingMoreMessages && (
                <Skeleton className={styles.msgLoadingShimmer}>
                  {/* Assistant message placeholder */}
                  <div className={styles.msgShimmerAssistantRow}>
                    <SkeletonItem shape="circle" size={32} />
                    <div className={styles.msgShimmerLines}>
                      <SkeletonItem shape="rectangle" size={16} className={styles.msgShimmerLineSm}/>
                      <SkeletonItem shape="rectangle" size={16} className={styles.msgShimmerLineMd} />
                      <SkeletonItem shape="rectangle" size={16} className={styles.msgShimmerLineSm} />
                    </div>
                  </div>
                  {/* User message placeholder — only shown when the current top message
                      is an assistant message (the turn above it was a user turn) */}
                  {messages[0]?.role === 'assistant' && (
                    <div className={styles.msgShimmerUserRow}>
                      <SkeletonItem shape="rectangle" size={16} className={styles.msgShimmerUserBubble} />
                    </div>
                  )}
                </Skeleton>
              )}
            </div>
          )}
          {messages.length === 0 && selectedAgent ? (
            <>
              {error && (
                <MessageBar intent="warning" className={styles.emptyStateNotice}>
                  <MessageBarBody>{error}</MessageBarBody>
                </MessageBar>
              )}
              <StarterMessages
                agent={selectedAgent}
                userLastName={userLastName}
                onPromptClick={onStarterPromptClick}
              />
            </>
          ) : messages.length === 0 ? (
            <div className={styles.loadingPlaceholder}>{strings.Loading}</div>
          ) : (
            <>
              {messages.map((msg, i) =>
                msg.role === 'user' ? (
                  <UserMessage key={i} message={msg} />
                ) : (
                  <AssistantMessage
                    key={i}
                    message={msg}
                    agentName={selectedAgent?.displayName ?? agents[0]?.displayName ?? strings.AgentDefaultName}
                    agentIcon={<AgentIcon agent={selectedAgent ?? agents[0]!} size={25} />}
                    messageIndex={i}
                    isStreaming={isStreaming && i === messages.length - 1}
                    showToolLoader={
                      i === messages.length - 1 && streamState === 'executing'
                    }
                    toolName={i === messages.length - 1 ? toolName : undefined}
                    activity={
                      (msg.steps && msg.steps.length > 0) || (isStreaming && i === messages.length - 1) ? (
                        <AgentActivity
                          steps={msg.steps ?? []}
                          streaming={isStreaming && i === messages.length - 1}
                          hasAnswer={!!msg.content}
                          startedAt={msg.startedAt}
                          endedAt={msg.endedAt}
                        />
                      ) : undefined
                    }
                    onGetDetails={
                      msg.id && onGetMessageDetails
                        ? () => onGetMessageDetails(msg.id!)
                        : undefined
                    }
                  />
                )
              ).map((el, i) => {
                const extras = messages[i]?.role === 'assistant' ? renderMessageExtras?.(messages[i], i) : null;
                return extras ? <React.Fragment key={`m${i}`}>{el}{extras}</React.Fragment> : <React.Fragment key={`m${i}`}>{el}</React.Fragment>;
              })}
              {streamState === 'executing' &&
                messages.length > 0 &&
                messages[messages.length - 1]?.role === 'user' && (
                  <AssistantMessage
                    key="tool-loader"
                    message={{ role: 'assistant', content: '' }}
                    agentName={selectedAgent?.displayName ?? agents[0]?.displayName ?? strings.AgentDefaultName}
                    agentIcon={<AgentIcon agent={selectedAgent ?? agents[0]!} size={25} />}
                    messageIndex={messages.length}
                    showToolLoader
                    toolName={toolName}
                  />
                )}
              {error && (
                <AssistantMessage
                  key="error-response"
                  message={{ role: 'assistant', content: '' }}
                  agentName={selectedAgent?.displayName ?? agents[0]?.displayName ?? strings.AgentDefaultName}
                  agentIcon={<AgentIcon agent={selectedAgent ?? agents[0]!} size={25} />}
                  messageIndex={messages.length}
                  isError={error}
                  onRetry={onRetry}
                />
              )}
              <div ref={messagesEndRef} />
            </>
          )}
        </div>
        </div>
      </div>
      <div className={styles.chatInputArea}>
        <div className={styles.jumpSlot} data-visible={showJump ? 'true' : undefined} aria-hidden={!showJump}>
          <Tooltip content={strings.ScrollToBottom} relationship="label">
            <Button
              className={styles.jumpButton}
              shape="circular"
              icon={<ArrowDown20Regular />}
              aria-label={strings.ScrollToLatestMessage}
              tabIndex={showJump ? 0 : -1}
              onClick={() => { scrollToBottom(true); focusInput(); }}
            />
          </Tooltip>
        </div>
        <div ref={inputWrapperRef} className={styles.inputWrapper}>
          {allowAttachments && (
            <input
              ref={fileInputRef}
              type="file"
              multiple
              hidden
              aria-hidden="true"
              onChange={handleFileChange}
            />
          )}
          {voiceActive && (
            <VoiceRecorderBar
              state={voice.state}
              elapsed={voice.elapsed}
              levels={voice.levels}
              onCancel={voice.cancel}
              onConfirm={() => { confirmVoice().catch(() => undefined); }}
            />
          )}
          <div style={voiceActive ? { display: 'none' } : undefined}>
          <CopilotChatInput
            aria-label={strings.ChatInputAriaLabel}
            charactersRemainingMessage={() => ''}
            isSending={isStreaming}
            disableSend={isUploading}
            onStop={isStreaming ? onCancelStream : () => { /* no-op while uploading */ }}
            maxLength={4000}
            onSubmit={handleSubmit}
            placeholderValue={strings.TypePlaceholder}
            actions={allowAttachments || (dictationLang && voice.supported) ? {
              children: (
                <>
                  {allowAttachments && (
                    <Button
                      appearance="subtle"
                      size="small"
                      shape="circular"
                      icon={<Attach24Regular />}
                      aria-label={strings.AttachFile}
                      onClick={() => fileInputRef.current?.click()}
                    />
                  )}
                  {dictationLang && voice.supported && (
                    <Tooltip content={voice.error ?? strings.DictateMessage} relationship="label">
                      <Button
                        appearance="subtle"
                        size="small"
                        shape="circular"
                        icon={<Mic24Regular />}
                        aria-label={strings.DictateMessage}
                        onClick={voice.start}
                        disabled={isStreaming}
                      />
                    </Tooltip>
                  )}
                </>
              )
            } : undefined}
            attachments={attachedFiles.length > 0 ? {
              children: (
                <AttachmentList onAttachmentDismiss={handleAttachmentDismiss}>
                  {attachedFiles.map((a) => (
                    <Tooltip
                      key={a.localId}
                      content={
                        a.error
                          ? (a.errorMessage ?? strings.FileUploadFailed.replace('{0}', a.file.name))
                          : a.file.name
                      }
                      relationship="label"
                      withArrow
                    >
                      <Attachment
                        id={a.localId}
                        dismissOnly
                        media={{
                          children: a.error
                            ? <ErrorCircle20Filled className={styles.attachmentErrorIcon} />
                            : getFileTypeIcon(a.file)
                        }}
                        content={{
                          children: a.error
                            ? strings.FileUploadFailed.replace('{0}', a.file.name)
                            : a.file.name
                        }}
                        progress={a.uploading ? { value: undefined, bar: { className: styles.attachmentProgressBar } } : undefined}
                      />
                    </Tooltip>
                  ))}
                </AttachmentList>
              )
            } : undefined}
          >
            <ImperativeControlPlugin ref={controlRef} />
          </CopilotChatInput>
          </div>
        </div>
      </div>
        </div>
        {rightPanel && <div className={styles.rightPanel}>{rightPanel}</div>}
        </div>
      </div>
      <Toaster toasterId={toasterId} position="bottom-end" />
      <Waves brandColor={wavesBrandColor} />
    </div>
  );
};
