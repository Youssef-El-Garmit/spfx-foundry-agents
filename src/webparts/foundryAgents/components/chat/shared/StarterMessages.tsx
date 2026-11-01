/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Subtitle1, Button } from '@fluentui/react-components';
import strings from 'FoundryAgentsWebPartStrings';
import type { IAgentConfig } from '../../../models';
import { AgentIcon } from './AgentIcon';
import { formatString } from '../../../utils/formatString';
import styles from './StarterMessages.module.scss';

const getFallbackStarterPrompts = (): string[] => [
  strings.DefaultSuggestionPrompt1,
  strings.DefaultSuggestionPrompt2,
  strings.DefaultSuggestionPrompt3
];

export interface IStarterMessagesProps {
  agent: IAgentConfig;
  userLastName?: string;
  onPromptClick?: (prompt: string) => void;
}

export const StarterMessages: React.FunctionComponent<IStarterMessagesProps> = ({
  agent,
  userLastName,
  onPromptClick
}) => {
  const welcomeText = userLastName
    ? formatString(strings.WelcomeMessage, userLastName)
    : strings.WelcomeMessageNoName;

  const prompts = agent.suggestionPrompts?.length ? agent.suggestionPrompts : getFallbackStarterPrompts();

  return (
    <div className={styles.zeroprompt}>
      <div className={styles.content}>
        <AgentIcon agent={agent} size={48} />
        <Subtitle1 className={styles.welcome}>{welcomeText}</Subtitle1>
      </div>

      {onPromptClick && prompts.length > 0 && (
        <ul className={styles.promptList}>
          {prompts.map((prompt, index) => (
            <li key={`prompt-${index}`}>
              <Button
                shape="circular"                
                onClick={() => onPromptClick(prompt)}
                title={prompt}
              >
                {prompt}
              </Button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};
