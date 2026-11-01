/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="./FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneCustomFieldProps,
  type IPropertyPaneField,
  PropertyPaneTextField,
  PropertyPaneLink,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import strings from 'FoundryAgentsWebPartStrings';
import FoundryAgents from './components/chat/FoundryAgents';
import type { IFoundryAgentsProps } from './components/chat/IFoundryAgentsProps';
import { getTeamsHostFromContext, type ITeamsAppContext } from './hooks/useTeamsHost';
import type { IAgentConfig } from './models';
import { AgentsPropertyPaneControl } from './propertyPane/AgentsPropertyPaneControl';
import { ExportImportPropertyPaneControl } from './propertyPane/ExportImportPropertyPaneControl';

// PropertyPaneCustomField exists at runtime but is excluded from public .d.ts
/* eslint-disable @typescript-eslint/no-var-requires */
const PropertyPaneCustomField = (require('@microsoft/sp-property-pane') as {
  PropertyPaneCustomField: (props: IPropertyPaneCustomFieldProps) => IPropertyPaneField<IPropertyPaneCustomFieldProps>;
}).PropertyPaneCustomField;
/* eslint-enable @typescript-eslint/no-var-requires */

/** Serialized in property pane: JSON array of IAgentConfig + appearance */
export interface IFoundryAgentsWebPartProps {
  /** JSON string of IAgentConfig[] */
  agentsJson: string;
  /** CSS min-height value for the web part container */
  minHeight: string;
}

function parseAgents(agentsJson: string | undefined): IAgentConfig[] {
  if (!agentsJson || !agentsJson.trim()) return [];
  try {
    const arr = JSON.parse(agentsJson) as unknown;
    if (!Array.isArray(arr)) return [];
    return arr.filter(
      (a): a is IAgentConfig =>
        a && typeof a === 'object' && typeof a.displayName === 'string' && typeof a.agentId === 'string' && typeof a.endpointUrl === 'string'
    );
  } catch {
    return [];
  }
}

export default class FoundryAgentsWebPart extends BaseClientSideWebPart<IFoundryAgentsWebPartProps> {
  private _requestOpenSettingsTrigger = 0;
  private _isPersonalApp: boolean | null = null;
  private _isTeamsTab: boolean | null = null;
  private _propertyPaneOpenOriginal: (() => void) | null = null;

  public render(): void {
    const agents = parseAgents(this.properties.agentsJson);
    const isEditMode = this.context.propertyPane.isPropertyPaneOpen?.() ?? false;
    const element: React.ReactElement<IFoundryAgentsProps> = React.createElement(FoundryAgents, {
      agents,
      isEditMode,
      onOpenPropertyPane: () => this._onOpenPropertyPaneRequested(),
      context: this.context,
      minHeight: this.properties.minHeight || 'min(600px, 90vh)',
      openSettingsTrigger: this._requestOpenSettingsTrigger || undefined
    });
    ReactDom.render(element, this.domElement);
  }

  private _onOpenPropertyPaneRequested(): void {
    if (this._isPersonalApp === true) {
      this._requestOpenSettingsTrigger = Date.now();
      this.render();
      return;
    }
    if (this._isPersonalApp === false && this._isTeamsTab === false) {
      this.context.propertyPane.open();
      return;
    }
    this._requestOpenSettingsTrigger = Date.now();
    this.render();
    this.context.propertyPane.open();
  }

  protected onInit(): Promise<void> {
    if (!this.properties.agentsJson) {
      this.properties.agentsJson = '[]';
    }
    if (!this.properties.minHeight) {
      this.properties.minHeight = 'min(600px, 90vh)';
    }
    this._propertyPaneOpenOriginal = this.context.propertyPane.open.bind(this.context.propertyPane);
    this.context.propertyPane.open = (): void => {
      if (this._isPersonalApp === true) {
        this._requestOpenSettingsTrigger = Date.now();
        this.render();
        return;
      }
      if (this._isPersonalApp === false && this._isTeamsTab === false) {
        this._propertyPaneOpenOriginal?.();
        return;
      }
      this._requestOpenSettingsTrigger = Date.now();
      this.render();
      this._propertyPaneOpenOriginal?.();
    };
    this._resolveTeamsContext();
    this._fillPageHeight();
    return Promise.resolve();
  }

  /**
   * SharePoint pages keep a 48px empty band under the canvas (spAppAndPropertyPanelContainer padding).
   * On a page hosting this chat, let the canvas use the full window height so the input sits at the bottom.
   */
  private _fillPageHeight(): void {
    if (typeof document === 'undefined' || document.getElementById('foundry-agents-fill-height')) return;
    if (!document.querySelector('.spAppAndPropertyPanelContainer')) return;
    const style = document.createElement('style');
    style.id = 'foundry-agents-fill-height';
    style.textContent = '.spAppAndPropertyPanelContainer{padding-bottom:0!important;height:calc(100% - 48px)!important}';
    document.head.appendChild(style);
  }

  protected onDispose(): void {
    document.getElementById('foundry-agents-fill-height')?.remove();
    ReactDom.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _resolveTeamsContext(): void {
    const ctx = this.context as { sdks?: { microsoftTeams?: { teamsJs?: { app?: { getContext(): Promise<ITeamsAppContext> } } } }; sdksAsync?: Promise<unknown> };
    const teamsJs = ctx.sdks?.microsoftTeams?.teamsJs;
    if (teamsJs?.app) {
      teamsJs.app.getContext().then((teamsContext) => {
        const { isPersonalApp, isTeamsTab } = getTeamsHostFromContext(teamsContext);
        this._isPersonalApp = isPersonalApp;
        this._isTeamsTab = isTeamsTab;
        this.render();
      }).catch(() => {
        this._isPersonalApp = false;
        this._isTeamsTab = false;
      });
      return;
    }
    const sdksAsync = ctx.sdksAsync;
    if (sdksAsync) {
      sdksAsync.then((resolved) => {
        const teams = (resolved as { microsoftTeams?: { teamsJs?: { app?: { getContext(): Promise<ITeamsAppContext> } } } })?.microsoftTeams;
        if (teams?.teamsJs?.app) {
          return teams.teamsJs.app.getContext();
        }
        return Promise.resolve<ITeamsAppContext | undefined>(undefined);
      }).then((teamsContext) => {
        const { isPersonalApp, isTeamsTab } = getTeamsHostFromContext(teamsContext);
        this._isPersonalApp = isPersonalApp;
        this._isTeamsTab = isTeamsTab;
        this.render();
      }).catch(() => {
        this._isPersonalApp = false;
        this._isTeamsTab = false;
      });
    }
  }

  protected onThemeChanged(_currentTheme: IReadonlyTheme | undefined): void {
    // Theme can be passed to FluentProvider when needed
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    /* eslint-disable @typescript-eslint/no-this-alias */
    const webPart = this;
    /* eslint-enable @typescript-eslint/no-this-alias */
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.AgentsGroupName,
              groupFields: [
                PropertyPaneCustomField({
                  key: 'agentsJson',
                  onRender: (
                    elem: HTMLElement,
                    _context: unknown,
                    onChanged: (targetProperty: string, value: unknown, isValid?: boolean) => void
                  ) => {
                    const agentsJson = webPart.properties.agentsJson || '';
                    const el = React.createElement(AgentsPropertyPaneControl, {
                      agentsJson,
                      onChange: (newJson: string) => onChanged('agentsJson', newJson, true),
                      context: webPart.context
                    });
                    ReactDom.render(el, elem);
                  },
                  onDispose: (elem: HTMLElement) => {
                    ReactDom.unmountComponentAtNode(elem);
                  }
                })
              ]
            },
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneTextField('minHeight', {
                  label: strings.MinHeightLabel,
                  description: strings.MinHeightDescription,
                  placeholder: 'min(600px, 90vh)'
                })
              ]
            },
            {
              groupName: strings.ExportImportSectionTitle,
              groupFields: [
                PropertyPaneCustomField({
                  key: 'exportImportAgentsJson',
                  onRender: (
                    elem: HTMLElement,
                    _context: unknown,
                    onChanged: (targetProperty: string, value: unknown, isValid?: boolean) => void
                  ) => {
                    const agentsJson = webPart.properties.agentsJson || '';
                    const el = React.createElement(ExportImportPropertyPaneControl, {
                      agentsJson,
                      onChange: (newJson: string) => onChanged('agentsJson', newJson, true)
                    });
                    ReactDom.render(el, elem);
                  },
                  onDispose: (elem: HTMLElement) => {
                    ReactDom.unmountComponentAtNode(elem);
                  }
                })
              ]
            },
            {
              groupName: strings.FooterGroupName,
              groupFields: [
                PropertyPaneLink('', {
                  text: strings.FooterWebsiteLink,
                  href: 'https://nemoura.com',
                  target: '_blank'
                }),
                PropertyPaneLink('', {
                  text: strings.FooterCompanyLinkedIn,
                  href: 'https://www.linkedin.com/company/n%C3%A9moura-communication/',
                  target: '_blank'
                }),
                PropertyPaneLink('', {
                  text: strings.FooterCompanyGitHub,
                  href: 'https://github.com/Nemoura-Labs',
                  target: '_blank'
                }),
                PropertyPaneLink('', {
                  text: strings.FooterDeveloperLinkedIn,
                  href: 'https://www.linkedin.com/in/youssef-elgarmit-spfx/',
                  target: '_blank'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
