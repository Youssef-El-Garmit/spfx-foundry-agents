/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import {
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Button,
  Input,
  Textarea,
  Radio,
  RadioGroup,
  Field,
  Switch,
  Checkbox,
  tokens
} from '@fluentui/react-components';
import strings from 'FoundryAgentsWebPartStrings';
import type { IAgentConfig } from '../../models';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { validateAgentInContext, type IAgentFieldErrors } from '../../utils/agentValidation';
import { FluentIconPicker } from './FluentIconPicker';
import { CustomIconInput } from './CustomIconInput';
import { EditablePromptList } from './EditablePromptList';
import { AudiencePicker } from './AudiencePicker';
import styles from './AgentEditDialog.module.scss';

export interface IAgentEditDialogProps {
  open: boolean;
  agent: IAgentConfig;
  onOpenChange: (open: boolean) => void;
  onSave: (agent: IAgentConfig) => void;
  /** Portal target so dialog renders inside FluentProvider and inherits theme */
  mountNode?: HTMLElement;
  /**
   * All other agents' display names (i.e. every agent except the one being edited).
   * Used to detect duplicate display names at save time.
   */
  otherAgents?: IAgentConfig[];
  /** SPFx context required by AudiencePicker to search Graph and SP groups. Optional: when absent the audience section is hidden. */
  context?: WebPartContext;
  /** When true, hide the audience section (e.g. for personal app where audience is not relevant). */
  hideAudience?: boolean;
}

const DEFAULT_BLOCKED_METHODS = ['DELETE', 'PATCH'];
/** All HTTP methods that can be blocked (RFC 7231 + PATCH). */
const BLOCKABLE_METHODS = ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'HEAD', 'OPTIONS'] as const;

function getDefaultSuggestionPrompts(): [string, string, string] {
  return [strings.DefaultSuggestionPrompt1, strings.DefaultSuggestionPrompt2, strings.DefaultSuggestionPrompt3];
}

/** Default additional context template using all variables for new agents. */
function getDefaultAdditionalInformationTemplate(): string {
  return [
    'Current user: {userDisplayName}',
    'Email: {userMail}',
    'User principal: {userPrincipalName}',
    'Date/time: {currentDateTime}',
    'Timezone: {currentTimezone}',
    'Region: {currentRegion}'
  ].join('\n');
}

function getBlockedMethods(agent: IAgentConfig): string[] {
  const raw = agent.blockedHttpMethods;
  if (!Array.isArray(raw)) return [...DEFAULT_BLOCKED_METHODS];
  if (raw.length === 0) return [];
  return raw.filter((m): m is string => typeof m === 'string' && m.trim() !== '');
}

/** Maps AgentFieldError codes to localized display strings. */
function toErrorMessage(code: IAgentFieldErrors[keyof IAgentFieldErrors]): string | undefined {
  if (!code) return undefined;
  if (code === 'invalid_url') return strings.InvalidUrlError;
  if (code === 'duplicate_name') return strings.DuplicateDisplayNameError;
  return strings.RequiredFieldError;
}

export const AgentEditDialog: React.FunctionComponent<IAgentEditDialogProps> = ({
  open,
  agent: initialAgent,
  onOpenChange,
  onSave,
  mountNode: mountNodeProp,
  otherAgents = [],
  context,
  hideAudience = false
}) => {
  const [agent, setAgent] = React.useState<IAgentConfig>(() => ({
    ...initialAgent,
    suggestionPrompts: initialAgent.suggestionPrompts?.length
      ? [...initialAgent.suggestionPrompts]
      : [...getDefaultSuggestionPrompts()],
    additionalInformationTemplate:
      initialAgent.additionalInformationTemplate?.trim() || getDefaultAdditionalInformationTemplate()
  }));

  const [errors, setErrors] = React.useState<IAgentFieldErrors>({});

  React.useEffect(() => {
    if (open) {
      setAgent({
        ...initialAgent,
        suggestionPrompts: initialAgent.suggestionPrompts?.length
          ? [...initialAgent.suggestionPrompts]
          : [...getDefaultSuggestionPrompts()],
        additionalInformationTemplate:
          initialAgent.additionalInformationTemplate?.trim() || getDefaultAdditionalInformationTemplate()
      });
      setErrors({});
    }
  }, [open, initialAgent]);

  const update = React.useCallback((patch: Partial<IAgentConfig>) => {
    setAgent((prev) => ({ ...prev, ...patch }));
    setErrors((prev) => {
      const next = { ...prev };
      if ('displayName' in patch) delete next.displayName;
      if ('agentId' in patch) delete next.agentId;
      if ('endpointUrl' in patch) delete next.endpointUrl;
      return next;
    });
  }, []);

  const handleSave = React.useCallback(() => {
    const fieldErrors = validateAgentInContext(agent, otherAgents);
    if (Object.keys(fieldErrors).length > 0) {
      setErrors(fieldErrors);
      return;
    }
    onSave(agent);
    onOpenChange(false);
  }, [agent, otherAgents, onSave, onOpenChange]);

  const boldLabelStyle: React.CSSProperties = {
    fontWeight: tokens.fontWeightSemibold as React.CSSProperties['fontWeight']
  };

  return (
    <Dialog open={open} onOpenChange={(_e, data) => onOpenChange(data.open)}>
      <DialogSurface
        mountNode={mountNodeProp}
        backdrop={{
          style: {
            zIndex:99999
          }
        }}
        style={{ maxWidth: 680, minWidth: 360, width: 'min(680px, calc(100vw - 32px))', zIndex:99999 }}
      >
        <DialogBody>
          <DialogTitle role="heading">
            <span
              style={{
                fontSize: tokens.fontSizeBase500,
                fontWeight: tokens.fontWeightBold as React.CSSProperties['fontWeight'],
                lineHeight: tokens.lineHeightBase500,
                color: tokens.colorNeutralForeground1
              }}
            >
              {strings.EditAgentDialogTitle}
            </span>
          </DialogTitle>
          <DialogContent>
            <div className={styles.formGrid}>
              {/* Always visible: core identity */}
              <Field
                required
                label={{ children: strings.AgentDisplayNameLabel, style: boldLabelStyle }}
                validationState={errors.displayName ? 'error' : undefined}
                validationMessage={toErrorMessage(errors.displayName)}
              >
                <Input
                  value={agent.displayName ?? ''}
                  onChange={(_e, data) => update({ displayName: data.value ?? '' })}
                  placeholder={strings.MyAgentPlaceholder}
                />
              </Field>

              <Field
                required
                label={{ children: strings.AgentIdLabel, style: boldLabelStyle }}
                validationState={errors.agentId ? 'error' : undefined}
                validationMessage={toErrorMessage(errors.agentId)}
              >
                <Input
                  value={agent.agentId ?? ''}
                  onChange={(_e, data) => update({ agentId: data.value ?? '' })}
                  placeholder={strings.AgentIdPlaceholder}
                />
              </Field>

              <div className={styles.fullSpan}>
                <Field
                  required
                  label={{ children: strings.EndpointUrlLabel, style: boldLabelStyle }}
                  hint={!errors.endpointUrl ? strings.EndpointUrlDescription : undefined}
                  validationState={errors.endpointUrl ? 'error' : undefined}
                  validationMessage={toErrorMessage(errors.endpointUrl)}
                >
                  <Input
                    value={agent.endpointUrl ?? ''}
                    onChange={(_e, data) => update({ endpointUrl: data.value ?? '' })}
                    placeholder={strings.EndpointUrlPlaceholder}
                  />
                </Field>
              </div>
            </div>

            <Accordion className={styles.accordion} collapsible multiple defaultOpenItems={[]}>
              {context && !hideAudience && (
                <AccordionItem value="audience" className={styles.accordionItem}>
                  <AccordionHeader className={styles.accordionHeader} expandIconPosition="end">{strings.AudienceGroupsLabel}</AccordionHeader>
                  <AccordionPanel>
                    <div className={styles.accordionPanelContent}>
                      <Field hint={strings.AudienceGroupsHint}>
                        <AudiencePicker
                          value={agent.audienceGroups ?? []}
                          onChange={(groups) => update({ audienceGroups: groups })}
                          context={context}
                          placeholder={strings.AudienceGroupsPlaceholder}
                        />
                      </Field>
                    </div>
                  </AccordionPanel>
                </AccordionItem>
              )}

              <AccordionItem value="appearance" className={styles.accordionItem}>
                <AccordionHeader className={styles.accordionHeader} expandIconPosition="end">{strings.IconAndConversationSectionLabel}</AccordionHeader>
                <AccordionPanel>
                  <div className={styles.accordionPanelContent}>
                    <Field label={{ children: strings.IconTypeLabel, style: boldLabelStyle }}>
                      <div className={styles.iconRow}>
                        <RadioGroup
                          value={agent.iconType}
                          onChange={(_e, data) => update({ iconType: data.value as 'fluent' | 'custom' })}
                          layout="horizontal"
                        >
                          <Radio value="fluent" label={strings.IconTypeFluent} />
                          <Radio value="custom" label={strings.IconTypeCustom} />
                        </RadioGroup>
                        <div className={styles.iconPickerWrap}>
                          {agent.iconType === 'fluent' ? (
                            <FluentIconPicker
                              value={agent.fluentIconName ?? 'Bot24Regular'}
                              onChange={(v) => update({ fluentIconName: v })}
                            />
                          ) : (
                            <CustomIconInput
                              value={agent.customIconUrl ?? ''}
                              onChange={(v) => update({ customIconUrl: v })}
                            />
                          )}
                        </div>
                      </div>
                    </Field>
                    <Field label={{ children: strings.HistoryFilterLabel, style: boldLabelStyle }} hint={strings.HistoryFilterHint}>
                      <Switch
                        checked={agent.showOnlyWebPartConversations === true}
                        onChange={(_e, data) => update({ showOnlyWebPartConversations: data.checked })}
                        label={agent.showOnlyWebPartConversations ? strings.HistoryFilterOnlyThis : strings.HistoryFilterAll}
                      />
                    </Field>
                    <Field label={{ children: strings.AllowAttachmentsLabel, style: boldLabelStyle }} hint={strings.AllowAttachmentsHint}>
                      <Switch
                        checked={agent.allowAttachments === true}
                        onChange={(_e, data) => update({ allowAttachments: data.checked })}
                        label={agent.allowAttachments ? strings.AllowAttachmentsOn : strings.AllowAttachmentsOff}
                      />
                    </Field>
                  </div>
                </AccordionPanel>
              </AccordionItem>

              <AccordionItem value="additionalContext" className={styles.accordionItem}>
                <AccordionHeader className={styles.accordionHeader} expandIconPosition="end">{strings.AdditionalInformationLabel}</AccordionHeader>
                <AccordionPanel>
                  <div className={styles.accordionPanelContent}>
                    <Field hint={strings.AdditionalInformationHint}>
                      <Textarea
                        value={agent.additionalInformationTemplate ?? ''}
                        onChange={(_e, data) => update({ additionalInformationTemplate: data.value ?? '' })}
                        placeholder={strings.AdditionalInformationPlaceholder}
                        rows={3}
                        resize="vertical"
                      />
                    </Field>
                  </div>
                </AccordionPanel>
              </AccordionItem>

              <AccordionItem value="voice" className={styles.accordionItem}>
                <AccordionHeader className={styles.accordionHeader} expandIconPosition="end">{strings.VoiceSectionLabel}</AccordionHeader>
                <AccordionPanel>
                  <div className={styles.accordionPanelContent}>
                    <Field label={{ children: strings.TranscriptionUrlLabel, style: boldLabelStyle }} hint={strings.TranscriptionUrlHint}>
                      <Input
                        value={agent.transcriptionUrl ?? ''}
                        onChange={(_e, data) => update({ transcriptionUrl: data.value ?? '' })}
                        placeholder={strings.TranscriptionUrlPlaceholder}
                      />
                    </Field>
                    <Field label={{ children: strings.TranscriptionPromptLabel, style: boldLabelStyle }} hint={strings.TranscriptionPromptHint}>
                      <Textarea
                        value={agent.transcriptionPrompt ?? ''}
                        onChange={(_e, data) => update({ transcriptionPrompt: data.value ?? '' })}
                        placeholder={strings.TranscriptionPromptPlaceholder}
                        rows={2}
                        resize="vertical"
                      />
                    </Field>
                  </div>
                </AccordionPanel>
              </AccordionItem>

              <AccordionItem value="security" className={styles.accordionItem}>
                <AccordionHeader className={styles.accordionHeader} expandIconPosition="end">{strings.SecuritySectionLabel}</AccordionHeader>
                <AccordionPanel>
                  <div className={styles.accordionPanelContent}>
                    <Field hint={strings.AllowFunctionCallsHint}>
                      <Switch
                        checked={agent.allowFunctionCalls === true}
                        onChange={(_e, data) => update({ allowFunctionCalls: data.checked })}
                        label={strings.AllowFunctionCallsLabel}
                      />
                    </Field>
                    {agent.allowFunctionCalls === true && (
                      <Field label={{ children: strings.BlockedHttpMethodsLabel, style: boldLabelStyle }} hint={strings.BlockedHttpMethodsHint}>
                        <div className={styles.checkboxRow}>
                          {BLOCKABLE_METHODS.map((method) => {
                            const blocked = getBlockedMethods(agent);
                            const checked = blocked.includes(method);
                            return (
                              <Checkbox
                                key={method}
                                id={`blocked-http-method-${method}`}
                                checked={checked}
                                onChange={(_e, data) => {
                                  const next = data.checked
                                    ? [...blocked.filter((m) => m !== method), method]
                                    : blocked.filter((m) => m !== method);
                                  update({ blockedHttpMethods: next });
                                }}
                                label={method}
                              />
                            );
                          })}
                        </div>
                      </Field>
                    )}
                  </div>
                </AccordionPanel>
              </AccordionItem>

              <AccordionItem value="suggestions" className={styles.accordionItem}>
                <AccordionHeader className={styles.accordionHeader} expandIconPosition="end">{strings.SuggestionPromptsLabel}</AccordionHeader>
                <AccordionPanel>
                  <div className={styles.accordionPanelContent}>
                    <EditablePromptList
                      prompts={agent.suggestionPrompts ?? []}
                      onChange={(prompts) => update({ suggestionPrompts: prompts })}
                    />
                  </div>
                </AccordionPanel>
              </AccordionItem>
            </Accordion>
          </DialogContent>
          <DialogActions position="end" style={{ paddingTop: tokens.spacingVerticalL }}>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>
              {strings.Cancel}
            </Button>
            <Button appearance="primary" onClick={handleSave}>
              {strings.SaveAgents}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
