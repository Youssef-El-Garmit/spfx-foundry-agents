/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import {
  Button,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Textarea,
  tokens,
} from '@fluentui/react-components';
import { ArrowExport24Regular, ArrowImport24Regular } from '@fluentui/react-icons';
import strings from 'FoundryAgentsWebPartStrings';
import type { IAgentConfig } from '../../models';

function getDefaultSuggestionPrompts(): string[] {
  return [strings.DefaultSuggestionPrompt1, strings.DefaultSuggestionPrompt2, strings.DefaultSuggestionPrompt3];
}

function parseImportedAgents(json: string): IAgentConfig[] | undefined {
  if (!json?.trim()) return undefined;
  try {
    const arr = JSON.parse(json) as unknown;
    if (!Array.isArray(arr)) return undefined;
    const defaultPrompts = getDefaultSuggestionPrompts();
    return arr
      .filter(
        (a): a is IAgentConfig =>
          a &&
          typeof a === 'object' &&
          typeof (a as IAgentConfig).displayName === 'string' &&
          typeof (a as IAgentConfig).agentId === 'string' &&
          typeof (a as IAgentConfig).endpointUrl === 'string'
      )
      .map((a) => ({ ...a, suggestionPrompts: a.suggestionPrompts ?? [...defaultPrompts] }));
  } catch {
    return undefined;
  }
}

export interface IExportImportSettingsProps {
  agents: IAgentConfig[];
  onSave: (agents: IAgentConfig[]) => void;
  /** Portal target for Dialog */
  portalMountNode?: HTMLElement;
  /** When true (e.g. personal app panel), show divider and section label. When false (native property pane group), omit them. */
  showSectionStyle?: boolean;
}

/**
 * Export / Import block: two full-width buttons in column with gap.
 * Export downloads agents as JSON; Import opens dialog to paste JSON and apply.
 */
export const ExportImportSettings: React.FunctionComponent<IExportImportSettingsProps> = ({
  agents,
  onSave,
  portalMountNode,
  showSectionStyle = true,
}) => {
  const [importDialogOpen, setImportDialogOpen] = React.useState(false);
  const [importText, setImportText] = React.useState('');
  const [importError, setImportError] = React.useState<string | null>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleExport = React.useCallback(() => {
    const json = JSON.stringify(agents, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'foundry-agents-settings.json';
    a.click();
    URL.revokeObjectURL(url);
  }, [agents]);

  const handleImportOpen = React.useCallback(() => {
    setImportText('');
    setImportError(null);
    setImportDialogOpen(true);
  }, []);

  const handleImportApply = React.useCallback(() => {
    const parsed = parseImportedAgents(importText);
    if (parsed === undefined) {
      setImportError(strings.ImportErrorInvalid);
      return;
    }
    onSave(parsed);
    setImportError(null);
    setImportDialogOpen(false);
  }, [importText, onSave]);

  const handleChooseFile = React.useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleFileChange = React.useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const text = typeof reader.result === 'string' ? reader.result : '';
      setImportText(text);
      setImportError(null);
    };
    reader.readAsText(file, 'utf-8');
    e.target.value = '';
  }, []);

  return (
    <>
      <div
        style={{
          display: 'flex',
          flexDirection: 'column',
          gap: tokens.spacingVerticalM,
          width: '100%',
          ...(showSectionStyle
            ? {
                marginTop: tokens.spacingVerticalL,
                paddingTop: tokens.spacingVerticalL,
                borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
              }
            : {}),
        }}
      >
        {showSectionStyle && (
          <span
            style={{
              fontSize: tokens.fontSizeBase200,
              fontWeight: tokens.fontWeightSemibold,
              color: tokens.colorNeutralForeground2,
            }}
          >
            {strings.ExportImportSectionTitle}
          </span>
        )}
        <Button
          appearance="secondary"
          icon={<ArrowExport24Regular />}
          onClick={handleExport}
          style={{ width: '100%' }}
        >
          {strings.ExportSettingsButton}
        </Button>
        <Button
          appearance="secondary"
          icon={<ArrowImport24Regular />}
          onClick={handleImportOpen}
          style={{ width: '100%' }}
        >
          {strings.ImportSettingsButton}
        </Button>
      </div>

      <Dialog open={importDialogOpen} onOpenChange={(_e, data) => !data.open && setImportDialogOpen(false)}>
        <DialogSurface mountNode={portalMountNode} backdrop={{ style: { zIndex: 99999 } }} style={{ zIndex: 99999 }}>
          <DialogBody>
            <DialogTitle>{strings.ImportDialogTitle}</DialogTitle>
            <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
              <input
                ref={fileInputRef}
                type="file"
                accept=".json,application/json"
                style={{ display: 'none' }}
                onChange={handleFileChange}
              />
              <Button
                appearance="secondary"
                onClick={handleChooseFile}
                style={{ alignSelf: 'flex-start' }}
              >
                {strings.ImportFromFileButton}
              </Button>
              <Textarea
                value={importText}
                onChange={(_e, d) => {
                  setImportText(d.value);
                  setImportError(null);
                }}
                placeholder={strings.ImportPlaceholder}
                rows={8}
                style={{ width: '100%', minHeight: 120 }}
              />
              {importError && (
                <span style={{ color: tokens.colorPaletteRedForeground1, fontSize: tokens.fontSizeBase200 }}>
                  {importError}
                </span>
              )}
            </DialogContent>
            <DialogActions position="end">
              <Button appearance="secondary" onClick={() => setImportDialogOpen(false)}>
                {strings.Cancel}
              </Button>
              <Button appearance="primary" onClick={handleImportApply}>
                {strings.ImportApplyButton}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
};
