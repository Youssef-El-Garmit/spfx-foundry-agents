/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Button, Tooltip } from '@fluentui/react-components';
import { Info16Regular } from '@fluentui/react-icons';
import strings from 'FoundryAgentsWebPartStrings';
import { Waves } from '../chat/shared/Waves';
import styles from './ConfigRequiredMessage.module.scss';

const WEBPART_LOGO_URL = 'https://nemoura.com/wp-content/uploads/assets/microsoft/SPFxFoundryAgents.png';

export interface IConfigRequiredMessageProps {
  onOpenPropertyPane: () => void;
  inline?: boolean;
  /** Principal message (e.g. admin or you should configure). Defaults to AdminMustConfigure */
  description?: string;
  /** When set (e.g. Teams channel tab), shown as tooltip at end of description; button is hidden */
  teamsTabConfigureHint?: string;
  /** When set, waves use this brand color for theme consistency */
  wavesBrandColor?: string;
}

export const ConfigRequiredMessage: React.FunctionComponent<IConfigRequiredMessageProps> = ({
  onOpenPropertyPane,
  description,
  teamsTabConfigureHint,
  wavesBrandColor
}) => {
  const mainDescription = description ?? strings.AdminMustConfigure;
  return (
    <div className={styles.root}>
      <div className={styles.content}>
        <img
          src={WEBPART_LOGO_URL}
          alt=""
          className={styles.logo}
          aria-hidden
        />
        <h2 className={styles.title}>{strings.ConfigNotSetTitle}</h2>
        <p className={styles.description}>
          {mainDescription}
          {teamsTabConfigureHint && (
            <>
              {' '}
              <Tooltip content={teamsTabConfigureHint} relationship="label" withArrow>
                <span className={styles.tooltipTrigger} tabIndex={0}>
                  <Info16Regular aria-hidden />
                </span>
              </Tooltip>
            </>
          )}
        </p>
        {!teamsTabConfigureHint && (
          <div className={styles.actions}>
            <Button appearance="primary" size="large" onClick={onOpenPropertyPane}>
              {strings.OpenWebPartProperties}
            </Button>
          </div>
        )}
      </div>
      <Waves brandColor={wavesBrandColor} />
    </div>
  );
};
