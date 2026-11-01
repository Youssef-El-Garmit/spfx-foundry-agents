/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import {
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Input,
  Label,
  Button,
  Radio,
  RadioGroup
} from '@fluentui/react-components';
import { Delete24Regular } from '@fluentui/react-icons';
import strings from 'FoundryAgentsWebPartStrings';
import type { IAgentConfig } from '../../models';
import { FluentIconPicker } from './FluentIconPicker';
import { CustomIconInput } from './CustomIconInput';
import { AgentIcon } from '../chat/shared/AgentIcon';

export interface IAgentConfigCardProps {
  agent: IAgentConfig;
  index: number;
  onUpdate: (agent: IAgentConfig) => void;
  onRemove: () => void;
  canRemove: boolean;
}

function defaultAgent(index: number, defaultSuggestionPrompts?: string[]): IAgentConfig {
  const suggestionPrompts = defaultSuggestionPrompts?.length
    ? [...defaultSuggestionPrompts]
    : [strings.DefaultSuggestionPrompt1, strings.DefaultSuggestionPrompt2, strings.DefaultSuggestionPrompt3];
  return {
    displayName: strings.AgentNumberTitle.replace('{0}', String(index + 1)),
    agentId: '',
    endpointUrl: '',
    iconType: 'fluent',
    fluentIconName: 'Bot24Regular',
    suggestionPrompts
  };
}

export const AgentConfigCard: React.FunctionComponent<IAgentConfigCardProps> = ({
  agent,
  index,
  onUpdate,
  onRemove,
  canRemove
}) => {
  const update = React.useCallback(
    (patch: Partial<IAgentConfig>) => {
      onUpdate({ ...agent, ...patch });
    },
    [agent, onUpdate]
  );

  const title = agent.displayName?.trim() || strings.AgentNumberTitle.replace('{0}', String(index + 1));

  return (
    <AccordionItem value={`agent-${index}`}>
      <AccordionHeader expandIconPosition="end">
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, width: '100%' }}>
          <AgentIcon agent={agent} size={24} />
          <span style={{ flex: 1 }}>{title}</span>
          {canRemove && (
            <Button
              appearance="subtle"
              icon={<Delete24Regular />}
              aria-label={strings.RemoveAgentButton}
              onClick={(e) => {
                e.stopPropagation();
                onRemove();
              }}
            />
          )}
        </div>
      </AccordionHeader>
      <AccordionPanel>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16, padding: '8px 0' }}>
          <div>
            <Label htmlFor={`displayName-${index}`}>{strings.AgentDisplayNameLabel}</Label>
            <Input
              id={`displayName-${index}`}
              value={agent.displayName ?? ''}
              onChange={(_e, data) => update({ displayName: data.value ?? '' })}
              placeholder={strings.MyAgentPlaceholder}
            />
          </div>
          <div>
            <Label htmlFor={`agentId-${index}`}>{strings.AgentIdLabel}</Label>
            <Input
              id={`agentId-${index}`}
              value={agent.agentId ?? ''}
              onChange={(_e, data) => update({ agentId: data.value ?? '' })}
              placeholder={strings.AgentIdPlaceholder}
            />
          </div>
          <div>
            <Label htmlFor={`endpointUrl-${index}`}>{strings.EndpointUrlLabel}</Label>
            <Input
              id={`endpointUrl-${index}`}
              value={agent.endpointUrl ?? ''}
              onChange={(_e, data) => update({ endpointUrl: data.value ?? '' })}
              placeholder={strings.EndpointUrlPlaceholder}
            />
            <span style={{ fontSize: 12, color: '#605e5c' }}>{strings.EndpointUrlDescription}</span>
          </div>
          <div>
            <Label htmlFor={`transcriptionUrl-${index}`}>{strings.TranscriptionUrlLabel}</Label>
            <Input
              id={`transcriptionUrl-${index}`}
              value={agent.transcriptionUrl ?? ''}
              onChange={(_e, data) => update({ transcriptionUrl: data.value ?? '' })}
              placeholder={strings.TranscriptionUrlPlaceholder}
            />
            <span style={{ fontSize: 12, color: '#605e5c' }}>{strings.TranscriptionUrlHint}</span>
          </div>
          <div>
            <Label htmlFor={`transcriptionPrompt-${index}`}>{strings.TranscriptionPromptLabel}</Label>
            <Input
              id={`transcriptionPrompt-${index}`}
              value={agent.transcriptionPrompt ?? ''}
              onChange={(_e, data) => update({ transcriptionPrompt: data.value ?? '' })}
              placeholder={strings.TranscriptionPromptPlaceholder}
            />
            <span style={{ fontSize: 12, color: '#605e5c' }}>{strings.TranscriptionPromptHint}</span>
          </div>
          <div>
            <Label>{strings.IconTypeLabel}</Label>
            <RadioGroup
              value={agent.iconType}
              onChange={(_e, data) => update({ iconType: data.value as 'fluent' | 'custom' })}
              layout="horizontal"
            >
              <Radio value="fluent" label={strings.IconTypeFluent} />
              <Radio value="custom" label={strings.IconTypeCustom} />
            </RadioGroup>
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
      </AccordionPanel>
    </AccordionItem>
  );
};

export { defaultAgent };
