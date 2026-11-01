/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Button, Input, tokens } from '@fluentui/react-components';
import { Delete24Regular, Edit24Regular, Add24Regular } from '@fluentui/react-icons';
import strings from 'FoundryAgentsWebPartStrings';
import { ListCardRow } from './ListCardRow';

export interface IEditablePromptListProps {
  prompts: string[];
  onChange: (prompts: string[]) => void;
  /** Section title (e.g. "Suggestion prompts"). Omit to not show a label (e.g. when parent already has a Field label) */
  title?: string;
  /** Label for "Add" button */
  addButtonLabel?: string;
  /** Placeholder for new prompt input */
  addPlaceholder?: string;
  /** Empty message when list is empty */
  emptyMessage?: string;
}

export const EditablePromptList: React.FunctionComponent<IEditablePromptListProps> = ({
  prompts,
  onChange,
  title,
  addButtonLabel = strings.AddSuggestionPromptButton,
  addPlaceholder = strings.SuggestionPromptPlaceholder,
  emptyMessage = strings.SuggestionPromptsEmpty
}) => {
  const [editingIndex, setEditingIndex] = React.useState<number | null>(null);
  const [editingValue, setEditingValue] = React.useState('');
  const [newPromptValue, setNewPromptValue] = React.useState('');
  const [dragOverIndex, setDragOverIndex] = React.useState<number | null>(null);

  const commitEdit = React.useCallback(() => {
    if (editingIndex === null) return;
    const trimmed = editingValue.trim();
    if (trimmed) {
      const next = [...prompts];
      next[editingIndex] = trimmed;
      onChange(next);
    } else {
      onChange(prompts.filter((_, i) => i !== editingIndex));
    }
    setEditingIndex(null);
    setEditingValue('');
  }, [editingIndex, editingValue, onChange, prompts]);

  const startEdit = React.useCallback((index: number) => {
    setEditingIndex(index);
    setEditingValue(prompts[index] ?? '');
  }, [prompts]);

  const handleKeyDown = React.useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit();
      }
      if (e.key === 'Escape') {
        setEditingIndex(null);
        setEditingValue('');
      }
    },
    [commitEdit]
  );

  const moveUp = React.useCallback(
    (index: number) => {
      if (index <= 0) return;
      const next = [...prompts];
      [next[index - 1], next[index]] = [next[index]!, next[index - 1]!];
      onChange(next);
    },
    [onChange, prompts]
  );

  const moveDown = React.useCallback(
    (index: number) => {
      if (index >= prompts.length - 1) return;
      const next = [...prompts];
      [next[index], next[index + 1]] = [next[index + 1]!, next[index]!];
      onChange(next);
    },
    [onChange, prompts]
  );

  const handleDragStart = React.useCallback((e: React.DragEvent, index: number) => {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', String(index));
    e.dataTransfer.setData('application/x-prompt-index', String(index));
  }, []);

  const handleDragOver = React.useCallback((e: React.DragEvent, index: number) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setDragOverIndex(index);
  }, []);

  const handleDragLeave = React.useCallback(() => {
    setDragOverIndex(null);
  }, []);

  const handleDrop = React.useCallback(
    (e: React.DragEvent, dropIndex: number) => {
      e.preventDefault();
      setDragOverIndex(null);
      const fromIndex = parseInt(e.dataTransfer.getData('application/x-prompt-index'), 10);
      if (Number.isNaN(fromIndex) || fromIndex === dropIndex) return;
      const next = [...prompts];
      const [removed] = next.splice(fromIndex, 1);
      next.splice(dropIndex, 0, removed);
      onChange(next);
    },
    [onChange, prompts]
  );

  const handleAdd = React.useCallback(() => {
    const trimmed = newPromptValue.trim();
    if (trimmed) {
      onChange([...prompts, trimmed]);
      setNewPromptValue('');
    }
  }, [onChange, prompts, newPromptValue]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS }}>
      {title && (
        <span style={{ fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground1 }}>
          {title}
        </span>
      )}
      <div
        style={{
          display: 'flex',
          flexDirection: 'column',
          gap: 10,
          maxHeight: 220,
          overflowY: 'auto',
        }}
      >
        {prompts.length === 0 ? (
          <div style={{ color: tokens.colorNeutralForeground3, fontSize: 12, padding: tokens.spacingVerticalS }}>
            {emptyMessage}
          </div>
        ) : (
          prompts.map((text, index) => (
            <div
              key={`${index}-${text}`}
              draggable
              onDragStart={(e) => handleDragStart(e, index)}
              onDragOver={(e) => handleDragOver(e, index)}
              onDragLeave={handleDragLeave}
              onDrop={(e) => handleDrop(e, index)}
              style={{
                cursor: 'grab',
                borderRadius: tokens.borderRadiusXLarge,
                outline: dragOverIndex === index ? `2px solid ${tokens.colorCompoundBrandStroke}` : undefined,
                margin: 2
              }}
            >
              <ListCardRow
                header={
                editingIndex === index ? (
                  <Input
                    value={editingValue}
                    onChange={(_e, data) => setEditingValue(data.value ?? '')}
                    onKeyDown={handleKeyDown}
                    onBlur={commitEdit}
                    placeholder={addPlaceholder}
                    autoFocus
                    style={{ minWidth: 180 }}
                  />
                ) : (
                  <span
                    style={{
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                      maxWidth: 280
                    }}
                    title={text}
                  >
                    {text || strings.SuggestionPromptPlaceholder}
                  </span>
                )
              }
              action={
                editingIndex !== index ? (
                  <div style={{ display: 'flex', gap: tokens.spacingHorizontalXXS }}>
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<Edit24Regular />}
                      aria-label={strings.EditSuggestionPrompt}
                      onClick={() => startEdit(index)}
                    />
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<Delete24Regular />}
                      aria-label={strings.DeleteSuggestionPrompt}
                      onClick={() => onChange(prompts.filter((_, i) => i !== index))}
                    />
                  </div>
                ) : null
              }
              onMoveUp={() => moveUp(index)}
              onMoveDown={() => moveDown(index)}
              canMoveUp={index > 0}
              canMoveDown={index < prompts.length - 1}
              style={dragOverIndex === index ? { backgroundColor: tokens.colorNeutralBackground3 } : undefined}
            />
            </div>
          ))
        )}
      </div>
      <div style={{ display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'center', flexWrap: 'wrap' }}>
        <Input
          value={newPromptValue}
          onChange={(_e, data) => setNewPromptValue(data.value ?? '')}
          onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), handleAdd())}
          placeholder={addPlaceholder}
          style={{ minWidth: 180, flex: 1 }}
        />
        <Button appearance="secondary" icon={<Add24Regular />} onClick={handleAdd} size="small" disabled={!newPromptValue.trim()}>
          {addButtonLabel}
        </Button>
      </div>
    </div>
  );
};
