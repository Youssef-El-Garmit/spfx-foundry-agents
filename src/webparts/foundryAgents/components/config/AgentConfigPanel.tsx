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
  tokens
} from '@fluentui/react-components';
import { Add24Regular, Edit24Regular, Delete24Regular } from '@fluentui/react-icons';
import strings from 'FoundryAgentsWebPartStrings';
import type { IAgentConfig } from '../../models';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { AgentIcon } from '../chat/shared/AgentIcon';
import { AgentEditDialog } from './AgentEditDialog';
import { defaultAgent } from './AgentConfigCard';
import { ListCardRow } from './ListCardRow';

function getDefaultSuggestionPrompts(): string[] {
  return [strings.DefaultSuggestionPrompt1, strings.DefaultSuggestionPrompt2, strings.DefaultSuggestionPrompt3];
}

export interface IAgentConfigPanelProps {
  agents: IAgentConfig[];
  onSave: (agents: IAgentConfig[]) => void;
  /** Portal target so Dialog/Tooltip render inside FluentProvider and inherit theme */
  portalMountNode?: HTMLElement;
  /** SPFx context forwarded to AgentEditDialog for audience group search */
  context?: WebPartContext;
  /** When true, hide the audience section in the edit dialog (e.g. personal app). */
  hideAudience?: boolean;
}

export const AgentConfigPanel: React.FunctionComponent<IAgentConfigPanelProps> = ({
  agents: initialAgents,
  onSave,
  portalMountNode,
  context,
  hideAudience = false
}) => {
  const [agents, setAgents] = React.useState<IAgentConfig[]>(() =>
    initialAgents.map((a) => ({ ...a, suggestionPrompts: a.suggestionPrompts ?? [...getDefaultSuggestionPrompts()] }))
  );

  const lastInitialRef = React.useRef<string>('');
  React.useEffect(() => {
    const serialized = JSON.stringify(initialAgents);
    if (serialized !== lastInitialRef.current) {
      lastInitialRef.current = serialized;
      setAgents(
        initialAgents.map((a) => ({ ...a, suggestionPrompts: a.suggestionPrompts ?? [...getDefaultSuggestionPrompts()] }))
      );
    }
  }, [initialAgents]);

  const [editingIndex, setEditingIndex] = React.useState<number | null>(null);
  const [deleteConfirmIndex, setDeleteConfirmIndex] = React.useState<number | null>(null);
  /** Holds the blank draft for a new agent while the "add" dialog is open. Null when closed. */
  const [newAgentDraft, setNewAgentDraft] = React.useState<IAgentConfig | null>(null);
  const [dragOverIndex, setDragOverIndex] = React.useState<number | null>(null);

  const handleMoveUp = React.useCallback((index: number) => {
    if (index <= 0) return;
    const next = [...agents];
    [next[index - 1], next[index]] = [next[index]!, next[index - 1]!];
    setAgents(next);
    onSave(next);
  }, [agents, onSave]);

  const handleMoveDown = React.useCallback((index: number) => {
    if (index >= agents.length - 1) return;
    const next = [...agents];
    [next[index], next[index + 1]] = [next[index + 1]!, next[index]!];
    setAgents(next);
    onSave(next);
  }, [agents, onSave]);

  const handleDragStart = React.useCallback((e: React.DragEvent, index: number) => {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', String(index));
    e.dataTransfer.setData('application/x-agent-index', String(index));
  }, []);

  const handleDragOver = React.useCallback((e: React.DragEvent, index: number) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setDragOverIndex(index);
  }, []);

  const handleDragLeave = React.useCallback(() => {
    setDragOverIndex(null);
  }, []);

  const handleDrop = React.useCallback((e: React.DragEvent, dropIndex: number) => {
    e.preventDefault();
    setDragOverIndex(null);
    const fromIndex = parseInt(e.dataTransfer.getData('application/x-agent-index'), 10);
    if (Number.isNaN(fromIndex) || fromIndex === dropIndex) return;
    const next = [...agents];
    const [removed] = next.splice(fromIndex, 1);
    next.splice(dropIndex, 0, removed);
    setAgents(next);
    onSave(next);
  }, [agents, onSave]);

  const handleUpdate = React.useCallback((index: number, updatedAgent: IAgentConfig) => {
    const next = agents.map((a, i) => (i === index ? updatedAgent : a));
    setAgents(next);
    onSave(next);
    setEditingIndex(null);
  }, [agents, onSave]);

  const handleCreate = React.useCallback((newAgent: IAgentConfig) => {
    const next = [...agents, newAgent];
    setAgents(next);
    onSave(next);
    setNewAgentDraft(null);
  }, [agents, onSave]);

  const handleRemove = React.useCallback((index: number) => {
    const next = agents.filter((_, i) => i !== index);
    setAgents(next);
    onSave(next);
    setDeleteConfirmIndex(null);
  }, [agents, onSave]);

  const handleAdd = React.useCallback(() => {
    setNewAgentDraft(defaultAgent(agents.length));
  }, [agents.length]);

  const title = (agent: IAgentConfig, index: number): string =>
    agent.displayName?.trim() || strings.AgentNumberTitle.replace('{0}', String(index + 1));

  return (
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalM
      }}
    >
      <div
        style={{
          display: 'flex',
          flexDirection: 'column',
          gap: tokens.spacingVerticalS,
          maxHeight: 300,
          overflowY: 'auto',
          paddingRight: tokens.spacingHorizontalXS
        }}
      >
        {agents.map((agent, index) => (
          <div
            key={index}
            draggable
            onDragStart={(e) => handleDragStart(e, index)}
            onDragOver={(e) => handleDragOver(e, index)}
            onDragLeave={handleDragLeave}
            onDrop={(e) => handleDrop(e, index)}
            style={{
              cursor: 'grab',
              outline: dragOverIndex === index ? `2px solid ${tokens.colorCompoundBrandStroke}` : undefined,
              borderRadius: tokens.borderRadiusXLarge,
              margin: 2
            }}
          >
            <ListCardRow
              image={<AgentIcon agent={agent} size={28} />}
              header={
                <span
                  style={{
                    fontWeight: tokens.fontWeightSemibold,
                    color: tokens.colorNeutralForeground1,
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                    minWidth: 0
                  }}
                >
                  {title(agent, index)}
                </span>
              }
              action={
                <div style={{ display: 'flex', gap: tokens.spacingHorizontalXXS }}>
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<Edit24Regular />}
                    aria-label={strings.EditAgentButton}
                    onClick={() => setEditingIndex(index)}
                  />
                  {agents.length > 1 && (
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<Delete24Regular />}
                      aria-label={strings.RemoveAgentButton}
                      onClick={() => setDeleteConfirmIndex(index)}
                    />
                  )}
                </div>
              }
              onMoveUp={() => handleMoveUp(index)}
              onMoveDown={() => handleMoveDown(index)}
              canMoveUp={index > 0}
              canMoveDown={index < agents.length - 1}
              style={dragOverIndex === index ? { backgroundColor: tokens.colorNeutralBackground3 } : undefined}
            />
          </div>
        ))}
      </div>

      <Button appearance="secondary" icon={<Add24Regular />} onClick={handleAdd}>
        {strings.AddAgentButton}
      </Button>

      {editingIndex !== null && agents[editingIndex] && (
        <AgentEditDialog
          open={true}
          agent={agents[editingIndex]}
          onOpenChange={(open) => !open && setEditingIndex(null)}
          onSave={(a) => handleUpdate(editingIndex, a)}
          mountNode={portalMountNode}
          otherAgents={agents.filter((_, i) => i !== editingIndex)}
          context={context}
          hideAudience={hideAudience}
        />
      )}
      {newAgentDraft && (
        <AgentEditDialog
          open={true}
          agent={newAgentDraft}
          onOpenChange={(open) => !open && setNewAgentDraft(null)}
          onSave={handleCreate}
          mountNode={portalMountNode}
          otherAgents={agents}
          context={context}
          hideAudience={hideAudience}
        />
      )}

      <Dialog open={deleteConfirmIndex !== null} onOpenChange={(_e, data) => !data.open && setDeleteConfirmIndex(null)}>
        <DialogSurface mountNode={portalMountNode} backdrop={{ style: { zIndex:99999 } } } style={{ zIndex:99999 }}>
          <DialogBody>
            <DialogTitle>{strings.DeleteConfirmTitle}</DialogTitle>
            <DialogContent>{strings.DeleteConfirmMessage}</DialogContent>
            <DialogActions position="end">
              <Button appearance="secondary" onClick={() => setDeleteConfirmIndex(null)}>
                {strings.Cancel}
              </Button>
              <Button
                appearance="primary"
                onClick={() => deleteConfirmIndex !== null && handleRemove(deleteConfirmIndex)}
              >
                {strings.RemoveAgentButton}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};
