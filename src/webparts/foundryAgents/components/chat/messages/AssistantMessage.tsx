/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { memo, useMemo, useCallback } from 'react';
import {
  Tooltip,
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Spinner,
  Divider,
} from '@fluentui/react-components';
import { CopilotMessage } from '@fluentui-copilot/react-copilot-chat';
import {
  DocumentRegular,
  GlobeRegular,
  FolderRegular,
  OpenRegular,
  Copy24Regular,
  CheckmarkCircle24Regular,
  MoreHorizontal24Regular,
} from '@fluentui/react-icons';
import copy from 'copy-to-clipboard';
import { format } from 'date-fns';
import { Markdown } from './Markdown';
import { ToolExecutionLoader } from '../shared/ToolExecutionLoader';
import { parseContentWithCitations } from '../../../utils/citationParser';
import type { IAnnotation } from '../../../types/chat';
import type { IChatMessage } from '../../../types/chat';
import type {
  IMessageDetailsData,
  IResponseOutputItem,
  IItemAnnotation,
  IConversationItemDetail,
} from '../../../services/FoundryAgentService';
import strings from 'FoundryAgentsWebPartStrings';
import styles from './AssistantMessage.module.scss';

export interface IAssistantMessageProps {
  message: IChatMessage;
  agentName?: string;
  /** Avatar element (e.g. <AgentIcon />) - required by CopilotMessage */
  agentIcon: React.ReactNode;
  messageIndex: number;
  isStreaming?: boolean;
  /** When set, renders an error state inside the message bubble instead of content. */
  isError?: string;
  /** Called when the user clicks Retry inside the error state. */
  onRetry?: () => void;
  /**
   * When provided, a "⋯" button is shown next to the copy button.
   * Clicking it calls this function to fetch full item + response details.
   * Only available when the message has a stable server-side id.
   */
  // eslint-disable-next-line @rushstack/no-new-null
  onGetDetails?: () => Promise<IMessageDetailsData | null>;
  /** When true, show the tool execution loader (thinking / running tool) inside this message. */
  showToolLoader?: boolean;
  /** Tool name for the loader label when showToolLoader is true. */
  toolName?: string;
  /** Live / past agent activity (thinking + tool actions) rendered above the answer. */
  activity?: React.ReactNode;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function formatTimestamp(isoString: string): string {
  try {
    const date = new Date(isoString);
    if (Number.isNaN(date.getTime())) return isoString;
    return format(date, 'PPp');
  } catch {
    return isoString;
  }
}

function formatUnixTimestamp(unix: number): string {
  try {
    return format(new Date(unix * 1000), 'PPp');
  } catch {
    return String(unix);
  }
}

// eslint-disable-next-line @rushstack/no-new-null
function extractReasoningText(output: IResponseOutputItem[] | undefined): string | null {
  if (!output) return null;
  for (const item of output) {
    if (item.type === 'reasoning') {
      if (typeof item.summary === 'string' && item.summary.trim()) return item.summary.trim();
      if (Array.isArray(item.summary)) {
        const text = item.summary.map(s => s.text ?? '').join('\n').trim();
        if (text) return text;
      }
    }
  }
  return null;
}

/** Collect unique annotations from the item's output_text content parts. */
function extractItemSources(item: IConversationItemDetail | undefined): IItemAnnotation[] {
  if (!item?.content) return [];
  const seen = new Set<string>();
  const result: IItemAnnotation[] = [];
  for (const part of item.content) {
    if (part.type !== 'output_text' || !part.annotations) continue;
    for (const ann of part.annotations) {
      const key = `${ann.type}:${ann.url ?? ann.file_id ?? ''}:${ann.title ?? ann.filename ?? ''}`;
      if (!seen.has(key)) {
        seen.add(key);
        result.push(ann);
      }
    }
  }
  return result;
}

function statusClass(status: string | undefined): string {
  if (status === 'completed') return styles.detailsStatusCompleted;
  if (status === 'failed') return styles.detailsStatusFailed;
  if (status === 'in_progress') return styles.detailsStatusInProgress;
  return '';
}

// ── Component ─────────────────────────────────────────────────────────────────

function AssistantMessageComponent({
  message,
  agentName = strings.AgentDefaultName,
  agentIcon,
  messageIndex,
  isStreaming = false,
  isError,
  onRetry,
  onGetDetails,
  showToolLoader: showToolLoaderProp = false,
  toolName,
  activity,
}: IAssistantMessageProps): React.ReactElement {
  // ── All hooks must be declared before any early return ────────────────────
  // The activity block replaces the generic loading dots / tool loader.
  const showToolLoader = showToolLoaderProp && !activity;
  const showLoadingDots = isStreaming && !message.content && !showToolLoader && !activity;
  const hasAnnotations =
    message.annotations !== undefined &&
    message.annotations !== null &&
    message.annotations.length > 0;

  const [copied, setCopied] = React.useState(false);
  const copyTimeoutRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  const [detailsOpen, setDetailsOpen] = React.useState(false);
  const [isLoadingDetails, setIsLoadingDetails] = React.useState(false);
  const [details, setDetails] = React.useState<IMessageDetailsData | null>(null);
  const [detailsError, setDetailsError] = React.useState<string | null>(null);
  const isFetchingRef = React.useRef(false);

  const handleCopyClick = useCallback(() => {
    copy(message.content);
    if (copyTimeoutRef.current) clearTimeout(copyTimeoutRef.current);
    setCopied(true);
    copyTimeoutRef.current = setTimeout(() => setCopied(false), 1500);
  }, [message.content]);

  const handleDetailsClick = useCallback((): void => {
    if (!onGetDetails) return;
    setDetailsOpen(true);
    if (!details && !isFetchingRef.current) {
      isFetchingRef.current = true;
      setIsLoadingDetails(true);
      setDetailsError(null);
      onGetDetails()
        .then(data => setDetails(data))
        .catch(e => setDetailsError(e instanceof Error ? e.message : strings.FailedToLoadDetails))
        .finally(() => {
          setIsLoadingDetails(false);
          isFetchingRef.current = false;
        });
    }
  }, [onGetDetails, details]);

  React.useEffect(
    () => () => {
      if (copyTimeoutRef.current) clearTimeout(copyTimeoutRef.current);
    },
    []
  );

  const parsedContent = useMemo(
    () => parseContentWithCitations(message.content, message.annotations),
    [message.content, message.annotations]
  );

  const indexedCitations = useMemo(() => {
    if (parsedContent.citations.length > 0) {
      return parsedContent.citations;
    }
    if (hasAnnotations) {
      const seen = new Map<string, { index: number; annotation: IAnnotation; count: number }>();
      message.annotations!.forEach(annotation => {
        const key = `${annotation.type}:${annotation.label}:${annotation.url ?? annotation.fileId ?? ''}`;
        const existing = seen.get(key);
        if (existing) {
          existing.count++;
        } else {
          seen.set(key, { index: seen.size + 1, annotation, count: 1 });
        }
      });
      return Array.from(seen.values());
    }
    return [];
  }, [parsedContent.citations, hasAnnotations, message.annotations]);

  const messageId = message.id ?? `msg-${messageIndex}`;

  const handleCitationClick = useCallback(
    (index: number, annotation?: IAnnotation) => {
      if (annotation?.type === 'uri_citation' && annotation.url) {
        window.open(annotation.url, '_blank', 'noopener,noreferrer');
      } else {
        const citationElement = document.getElementById(`citation-${messageId}-${index}`);
        if (citationElement) {
          citationElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
          citationElement.classList.add(styles.citationHighlight);
          setTimeout(() => citationElement.classList.remove(styles.citationHighlight), 2000);
        }
      }
    },
    [messageId]
  );

  const renderCitation = useCallback(
    (annotation: IAnnotation, index: number, count: number = 1): React.ReactElement => {
      const getIcon = (): React.ReactElement => {
        switch (annotation.type) {
          case 'uri_citation':
            return <GlobeRegular className={styles.citationIcon} />;
          case 'file_path':
            return <FolderRegular className={styles.citationIcon} />;
          default:
            return <DocumentRegular className={styles.citationIcon} />;
        }
      };

      const referencedStr = count > 1 ? ` (${strings.CitationReferencedTimes.replace('{0}', String(count))})` : '';
      const tooltipContent = annotation.quote
        ? `${annotation.label}${referencedStr}\n\n"${annotation.quote.slice(0, 200)}${annotation.quote.length > 200 ? '...' : ''}"`
        : `${annotation.label}${referencedStr}`;

      const isClickable = annotation.type === 'uri_citation' && !!annotation.url;

      const handleClick = (): void => {
        if (isClickable && annotation.url) {
          window.open(annotation.url, '_blank', 'noopener,noreferrer');
        }
      };

      const handleKeyDown = (e: React.KeyboardEvent): void => {
        if (isClickable && (e.key === 'Enter' || e.key === ' ')) {
          e.preventDefault();
          handleClick();
        }
      };

      return (
        <Tooltip
          key={`${annotation.label}-${index}`}
          content={tooltipContent}
          relationship="description"
          withArrow
        >
          <span
            id={`citation-${messageId}-${index}`}
            className={`${styles.citation} ${isClickable ? styles.citationClickable : ''}`}
            onClick={isClickable ? handleClick : undefined}
            onKeyDown={isClickable ? handleKeyDown : undefined}
            role={isClickable ? 'link' : undefined}
            tabIndex={isClickable ? 0 : undefined}
          >
            <span className={styles.citationNumber}>{index}</span>
            <span className={styles.citationContent}>
              {getIcon()}
              <span className={styles.citationLabel}>{annotation.label}</span>
              {count > 1 && <span className={styles.citationCount}>×{count}</span>}
              {isClickable && <OpenRegular className={styles.citationExternalIcon} />}
            </span>
          </span>
        </Tooltip>
      );
    },
    [messageId]
  );

  const citations = useMemo(
    () => indexedCitations.map(({ index, annotation, count }) => renderCitation(annotation, index, count)),
    [indexedCitations, renderCitation]
  );

  // ── Early return for error state ──────────────────────────────────────────
  if (isError) {
    return (
      <CopilotMessage
        id={`error-${messageIndex}`}
        avatar={agentIcon as React.ReactElement}
        name={agentName}
        loadingState="none"
        className={styles.copilotMessage}
      >
        <MessageBar intent="error" className={styles.errorBar}>
          <MessageBarBody>{isError}</MessageBarBody>
          {onRetry && (
            <MessageBarActions>
              <Button appearance="secondary" size="small" onClick={onRetry}>
                {strings.Retry}
              </Button>
            </MessageBarActions>
          )}
        </MessageBar>
      </CopilotMessage>
    );
  }

  // ── Derived values for the details dialog ─────────────────────────────────
  const reasoningText = extractReasoningText(details?.response?.output);
  const toolCalls = (details?.response?.output ?? []).filter(
    o =>
      o.type === 'file_search_call' ||
      o.type === 'web_search_call' ||
      o.type === 'function_call'
  );
  const responseStatus = details?.response?.status ?? details?.item?.status;
  const itemSources = extractItemSources(details?.item);

  const detailsDialog = onGetDetails ? (
    <Dialog open={detailsOpen} onOpenChange={(_, data) => setDetailsOpen(data.open)}>
      <DialogSurface className={styles.detailsSurface}>
        <DialogBody>
          <DialogTitle className={styles.detailsTitle}>{strings.ResponseDetailsTitle}</DialogTitle>
          <DialogContent className={styles.detailsContent}>
            {isLoadingDetails && (
              <div className={styles.detailsLoading}>
                <Spinner size="medium" label={strings.LoadingDetails} />
              </div>
            )}

            {detailsError && !isLoadingDetails && (
              <div style={{ padding: '12px 16px' }}>
                <MessageBar intent="error">
                  <MessageBarBody>{detailsError}</MessageBarBody>
                </MessageBar>
              </div>
            )}

            {details && !isLoadingDetails && (
              <>
                {/* ── Overview ─────────────────────────────────── */}
                <div className={styles.detailsSection}>
                  <div className={styles.detailsSectionTitle}>{strings.DetailsOverview}</div>

                  {details.response?.model && (
                    <div className={styles.detailsRow}>
                      <span className={styles.detailsLabel}>{strings.DetailsModel}</span>
                      <code className={styles.detailsCode}>{details.response.model}</code>
                    </div>
                  )}

                  {responseStatus && (
                    <div className={styles.detailsRow}>
                      <span className={styles.detailsLabel}>{strings.DetailsStatus}</span>
                      <span className={`${styles.detailsStatus} ${statusClass(responseStatus)}`}>
                        {responseStatus}
                      </span>
                    </div>
                  )}

                  {details.response?.created_at && (
                    <div className={styles.detailsRow}>
                      <span className={styles.detailsLabel}>{strings.DetailsCreated}</span>
                      <span>{formatUnixTimestamp(details.response.created_at)}</span>
                    </div>
                  )}

                  {details.response?.id && (
                    <div className={styles.detailsRow}>
                      <span className={styles.detailsLabel}>{strings.DetailsResponseId}</span>
                      <code className={styles.detailsCode}>{details.response.id}</code>
                    </div>
                  )}

                  {details.item.id && (
                    <div className={styles.detailsRow}>
                      <span className={styles.detailsLabel}>{strings.DetailsItemId}</span>
                      <code className={styles.detailsCode}>{details.item.id}</code>
                    </div>
                  )}
                </div>

                {/* ── Sources ──────────────────────────────────── */}
                {itemSources.length > 0 && (
                  <>
                    <Divider className={styles.detailsDivider} />
                    <div className={styles.detailsSection}>
                      <div className={styles.detailsSectionTitle}>
                        {strings.DetailsSourcesCount.replace('{0}', String(itemSources.length))}
                      </div>
                      <div className={styles.detailsSourceList}>
                        {itemSources.map((src, i) => {
                          const isUrl = src.type === 'url_citation' || src.type === 'uri_citation';
                          const isFile = src.type === 'file_citation' || src.type === 'file_path' || src.type === 'container_file_citation';
                          const label = src.title ?? src.filename ?? src.url ?? src.file_id ?? strings.DetailsSourceDefault.replace('{0}', String(i + 1));
                          return (
                            <div key={i} className={styles.detailsSourceItem}>
                              <div className={styles.detailsSourceHeader}>
                                {isUrl
                                  ? <GlobeRegular className={styles.detailsSourceIcon} />
                                  : <DocumentRegular className={styles.detailsSourceIcon} />
                                }
                                <span className={styles.detailsSourceLabel}>
                                  {isUrl && src.url ? (
                                    <a
                                      href={src.url}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className={styles.detailsSourceLink}
                                    >
                                      {label}
                                      <OpenRegular className={styles.detailsSourceExternalIcon} />
                                    </a>
                                  ) : (
                                    label
                                  )}
                                </span>
                                <span className={styles.detailsSourceTypeBadge}>
                                  {isUrl ? strings.DetailsWeb : isFile ? strings.DetailsFile : src.type}
                                </span>
                              </div>
                              {isUrl && src.url && src.url !== label && (
                                <div className={styles.detailsSourceUrl}>{src.url}</div>
                              )}
                              {isFile && src.file_id && (
                                <div className={styles.detailsSourceMeta}>
                                  <code className={styles.detailsCode}>{src.file_id}</code>
                                </div>
                              )}
                              {src.quote && (
                                <blockquote className={styles.detailsSourceQuote}>
                                  {src.quote.slice(0, 400)}{src.quote.length > 400 ? '…' : ''}
                                </blockquote>
                              )}
                              {src.text && !src.quote && (
                                <blockquote className={styles.detailsSourceQuote}>
                                  {src.text.slice(0, 400)}{src.text.length > 400 ? '…' : ''}
                                </blockquote>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </>
                )}

                {/* ── Reasoning ────────────────────────────────── */}
                {reasoningText && (
                  <>
                    <Divider className={styles.detailsDivider} />
                    <div className={styles.detailsSection}>
                      <div className={styles.detailsSectionTitle}>{strings.DetailsReasoning}</div>
                      <div className={styles.detailsReasoningText}>{reasoningText}</div>
                    </div>
                  </>
                )}

                {/* ── Tool Calls ───────────────────────────────── */}
                {toolCalls.length > 0 && (
                  <>
                    <Divider className={styles.detailsDivider} />
                    <div className={styles.detailsSection}>
                      <div className={styles.detailsSectionTitle}>{strings.DetailsToolCalls}</div>
                      {toolCalls.map((tc, i) => {
                        if (tc.type === 'file_search_call') {
                          return (
                            <div key={i} className={styles.detailsToolCall}>
                              <div className={styles.detailsToolCallHeader}>
                                <span>{strings.DetailsFileSearch}</span>
                                {tc.queries && tc.queries.length > 0 && (
                                  <span className={styles.detailsToolQuery}>
                                    &ldquo;{tc.queries.join('&rdquo;, &ldquo;')}&rdquo;
                                  </span>
                                )}
                              </div>
                              {tc.results && tc.results.length > 0 && (
                                <div className={styles.detailsToolResults}>
                                  {tc.results.map((r, j) => (
                                    <div key={j} className={styles.detailsToolResult}>
                                      <DocumentRegular className={styles.detailsToolResultIcon} />
                                      <div className={styles.detailsToolResultContent}>
                                        <div className={styles.detailsToolResultTitle}>
                                          {r.filename ?? r.file_id ?? strings.DetailsUnknownFile}
                                          {r.score !== undefined && (
                                            <span className={styles.detailsToolResultScore}>
                                              {(r.score * 100).toFixed(0)}%
                                            </span>
                                          )}
                                        </div>
                                        {r.text && (
                                          <div className={styles.detailsToolResultText}>
                                            &ldquo;{r.text.slice(0, 200)}{r.text.length > 200 ? '…' : ''}&rdquo;
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              )}
                            </div>
                          );
                        }

                        if (tc.type === 'web_search_call') {
                          return (
                            <div key={i} className={styles.detailsToolCall}>
                              <div className={styles.detailsToolCallHeader}>
                                <GlobeRegular className={styles.detailsToolResultIcon} />
                                <span>{strings.DetailsWebSearch}</span>
                                {tc.queries && tc.queries.length > 0 && (
                                  <span className={styles.detailsToolQuery}>
                                    &ldquo;{tc.queries.join('&rdquo;, &ldquo;')}&rdquo;
                                  </span>
                                )}
                              </div>
                              {tc.results && tc.results.length > 0 && (
                                <div className={styles.detailsToolResults}>
                                  {tc.results.map((r, j) => (
                                    <div key={j} className={styles.detailsToolResult}>
                                      <GlobeRegular className={styles.detailsToolResultIcon} />
                                      <div className={styles.detailsToolResultContent}>
                                        <div className={styles.detailsToolResultTitle}>
                                          {r.title ?? r.url ?? strings.DetailsResult}
                                        </div>
                                        {r.text && (
                                          <div className={styles.detailsToolResultText}>
                                            &ldquo;{r.text.slice(0, 200)}{r.text.length > 200 ? '…' : ''}&rdquo;
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              )}
                            </div>
                          );
                        }

                        if (tc.type === 'function_call') {
                          return (
                            <div key={i} className={styles.detailsToolCall}>
                              <div className={styles.detailsToolCallHeader}>
                                <span>{strings.DetailsFunction} <strong>{tc.name}</strong></span>
                              </div>
                              {tc.arguments && (
                                <pre className={styles.detailsFunctionArgs}>
                                  {(() => {
                                    try {
                                      return JSON.stringify(JSON.parse(tc.arguments), null, 2);
                                    } catch {
                                      return tc.arguments;
                                    }
                                  })()}
                                </pre>
                              )}
                            </div>
                          );
                        }

                        return null;
                      })}
                    </div>
                  </>
                )}

                {/* ── Token Usage ──────────────────────────────── */}
                {details.response?.usage && (
                  <>
                    <Divider className={styles.detailsDivider} />
                    <div className={styles.detailsSection}>
                      <div className={styles.detailsSectionTitle}>{strings.DetailsTokenUsage}</div>
                      <div className={styles.detailsUsageGrid}>
                        {details.response.usage.input_tokens !== undefined && (
                          <div className={styles.detailsUsageItem}>
                            <span className={styles.detailsUsageLabel}>{strings.DetailsInput}</span>
                            <span className={styles.detailsUsageValue}>
                              {details.response.usage.input_tokens.toLocaleString()}
                            </span>
                            {(details.response.usage.input_tokens_details?.cached_tokens ?? 0) > 0 && (
                              <span className={styles.detailsUsageSub}>
                                {details.response.usage.input_tokens_details!.cached_tokens!.toLocaleString()} {strings.DetailsCached}
                              </span>
                            )}
                          </div>
                        )}
                        {details.response.usage.output_tokens !== undefined && (
                          <div className={styles.detailsUsageItem}>
                            <span className={styles.detailsUsageLabel}>{strings.DetailsOutput}</span>
                            <span className={styles.detailsUsageValue}>
                              {details.response.usage.output_tokens.toLocaleString()}
                            </span>
                            {(details.response.usage.output_tokens_details?.reasoning_tokens ?? 0) > 0 && (
                              <span className={styles.detailsUsageSub}>
                                {details.response.usage.output_tokens_details!.reasoning_tokens!.toLocaleString()} {strings.DetailsReasoningTokens}
                              </span>
                            )}
                          </div>
                        )}
                        {details.response.usage.total_tokens !== undefined && (
                          <div className={`${styles.detailsUsageItem} ${styles.detailsUsageTotal}`}>
                            <span className={styles.detailsUsageLabel}>{strings.DetailsTotal}</span>
                            <span className={styles.detailsUsageValue}>
                              {details.response.usage.total_tokens.toLocaleString()}
                            </span>
                          </div>
                        )}
                      </div>
                    </div>
                  </>
                )}
              </>
            )}
          </DialogContent>

          <DialogActions>
            <Button appearance="primary" onClick={() => setDetailsOpen(false)}>
              {strings.DetailsClose}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  ) : null;

  // ── Footnote ──────────────────────────────────────────────────────────────
  const timestamp = message.more?.time ? formatTimestamp(message.more.time) : '';
  const usage = message.more?.usage;

  const footnote = (
    <div className={styles.footnoteContainer}>
      {indexedCitations.length > 0 && !isStreaming && (
        <div className={styles.citationList}>{citations}</div>
      )}
      <div className={styles.metadataRow}>
        {timestamp && <span className={styles.timestamp}>{timestamp}</span>}
        {usage && (
          <span className={styles.usageInfo}>
            {[
              usage.promptTokens !== undefined && usage.promptTokens !== null && `${usage.promptTokens} prompt`,
              usage.completionTokens !== undefined && usage.completionTokens !== null && `${usage.completionTokens} completion`,
            ]
              .filter(Boolean)
              .join(' · ')}
          </span>
        )}
        {message.content && (
          <Tooltip
            visible={copied}
            content={copied ? strings.Copied : ''}
            relationship="label"
            withArrow
          >
            <Button
              appearance="subtle"
              shape="circular"
              icon={copied ? <CheckmarkCircle24Regular /> : <Copy24Regular />}
              size="small"
              onClick={handleCopyClick}
              aria-label={copied ? strings.Copied : undefined}
              className={copied ? styles.copyButtonCopied : styles.copyButton}
            />
          </Tooltip>
        )}
        {onGetDetails && message.content && !isStreaming && (
          <Tooltip content={strings.ViewResponseDetails} relationship="label" withArrow>
            <Button
              appearance="subtle"
              shape="circular"
              icon={<MoreHorizontal24Regular />}
              size="small"
              onClick={handleDetailsClick}
              aria-label={strings.ViewResponseDetails}
              className={styles.detailsButton}
            />
          </Tooltip>
        )}
      </div>
      {detailsDialog}
    </div>
  );

  return (
    <CopilotMessage
      id={messageId}
      avatar={agentIcon as React.ReactElement}
      name={agentName}
      loadingState={showLoadingDots ? 'loading' : 'none'}
      className={styles.copilotMessage}
      footnote={footnote as React.ReactElement}
    >
      {activity}
      {showToolLoader && (
        <ToolExecutionLoader toolName={toolName} ariaLabel={strings.RunningTool} runningLabel={strings.RunningTool} toolFallback={strings.ToolGeneric} />
      )}
      {!showLoadingDots && !showToolLoader && !!message.content && (
        <Markdown
          content={message.content}
          annotations={message.annotations}
          onCitationClick={handleCitationClick}
        />
      )}
    </CopilotMessage>
  );
}

export const AssistantMessage = memo(AssistantMessageComponent);
