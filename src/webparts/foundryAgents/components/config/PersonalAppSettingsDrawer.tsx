/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import {
  Drawer,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerBody,
  DrawerFooter,
  Button,
  MessageBar,
  MessageBarBody,
  Tooltip,
  Link,
  tokens,
} from '@fluentui/react-components';
import { Dismiss24Regular, Info24Regular } from '@fluentui/react-icons';
import { AgentConfigPanel } from './AgentConfigPanel';
import { ExportImportSettings } from './ExportImportSettings';
import type { IAgentConfig } from '../../models';
import locStrings from 'FoundryAgentsWebPartStrings';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

function getErrorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (typeof err === 'object' && err !== null && 'message' in err) {
    return String((err as { message: unknown }).message);
  }
  return String(err);
}

export interface IPersonalAppSettingsDrawerStrings {
  PersonalAppSettingsLabel: string;
  DetailsClose: string;
}

export interface IPersonalAppSettingsDrawerProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  agents: IAgentConfig[];
  onSave: (agents: IAgentConfig[]) => void | Promise<void>;
  /** Error from initial extension load (shown in drawer when open). */
  extensionError: unknown;
  /** Error from last save attempt (shown in drawer when open). */
  saveError: unknown;
  portalMountNode?: HTMLElement;
  context?: WebPartContext;
  strings: IPersonalAppSettingsDrawerStrings;
}

/**
 * Drawer for personal app settings: loads/saves agents via Graph open extensions.
 * Displays any extension or save error as-is in the panel.
 */
export const PersonalAppSettingsDrawer: React.FunctionComponent<IPersonalAppSettingsDrawerProps> = ({
  open,
  onOpenChange,
  agents,
  onSave,
  extensionError,
  saveError,
  portalMountNode,
  context,
  strings,
}) => {
  const errorToShow = saveError ?? extensionError;
  const errorMessage = errorToShow !== null && errorToShow !== undefined ? getErrorMessage(errorToShow) : '';

  return (
    <Drawer
      type="overlay"
      open={open}
      onOpenChange={(_e, d) => onOpenChange(d.open)}
      position="end"
      size="medium"
      mountNode={portalMountNode ?? undefined}
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label={strings.DetailsClose}
              icon={<Dismiss24Regular />}
              onClick={() => onOpenChange(false)}
            />
          }
        >
          {strings.PersonalAppSettingsLabel}
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody style={{ overflowY: 'auto', display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', padding: 10 }}>
        <div style={{ width: '100%', maxWidth: '95%', flex: 1 }}>
          {errorToShow !== null && errorToShow !== undefined && errorMessage && (
            <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>
                {errorMessage}
                <Tooltip
                  content={
                    <span style={{ width: 'max-content', maxWidth: 'min(90vw, 600px)', display: 'inline-block' }}>
                      {errorMessage}
                    </span>
                  }
                  relationship="description"
                  withArrow
                >
                  <span style={{ marginLeft: 6, verticalAlign: 'middle', cursor: 'help' }} aria-label={errorMessage}>
                    <Info24Regular />
                  </span>
                </Tooltip>
              </MessageBarBody>
            </MessageBar>
          )}
          <AgentConfigPanel
            agents={agents}
            onSave={onSave}
            portalMountNode={portalMountNode}
            context={context}
            hideAudience={true}
          />
          <ExportImportSettings
            agents={agents}
            onSave={onSave}
            portalMountNode={portalMountNode}
          />
          <div
            style={{
              marginTop: tokens.spacingVerticalXL,
              paddingTop: tokens.spacingVerticalL,
              borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
              display: 'flex',
              flexDirection: 'column',
              gap: tokens.spacingVerticalS,
            }}
          >
            <span style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2 }}>
              {locStrings.FooterGroupName}
            </span>
            <Link href="https://nemoura.com" target="_blank" rel="noopener noreferrer" style={{ fontSize: tokens.fontSizeBase200 }}>
              {locStrings.FooterWebsiteLink}
            </Link>
            <Link href="https://www.linkedin.com/company/n%C3%A9moura-communication/" target="_blank" rel="noopener noreferrer" style={{ fontSize: tokens.fontSizeBase200 }}>
              {locStrings.FooterCompanyLinkedIn}
            </Link>
            <Link href="https://github.com/Nemoura-Labs" target="_blank" rel="noopener noreferrer" style={{ fontSize: tokens.fontSizeBase200 }}>
              {locStrings.FooterCompanyGitHub}
            </Link>
            <Link href="https://www.linkedin.com/in/youssef-elgarmit-spfx/" target="_blank" rel="noopener noreferrer" style={{ fontSize: tokens.fontSizeBase200 }}>
              {locStrings.FooterDeveloperLinkedIn}
            </Link>
          </div>
        </div>
      </DrawerBody>
      <DrawerFooter>
        <Button appearance="secondary" onClick={() => onOpenChange(false)}>
          {strings.DetailsClose}
        </Button>
      </DrawerFooter>
    </Drawer>
  );
};
