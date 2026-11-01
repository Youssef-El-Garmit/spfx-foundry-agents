/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Card, CardHeader, Button, tokens } from '@fluentui/react-components';
import { ChevronUp24Regular, ChevronDown24Regular } from '@fluentui/react-icons';
import strings from 'FoundryAgentsWebPartStrings';

export interface IListCardRowProps {
  /** Left slot: icon or other content (e.g. AgentIcon) */
  image?: React.ReactNode;
  /** Main content (title or text) */
  header: React.ReactNode;
  /** Right slot: action buttons (edit, delete) */
  action?: React.ReactNode;
  /** When set, show up/down order buttons to the left of image (stacked vertically) */
  onMoveUp?: () => void;
  /** When set, show up/down order buttons to the left of image */
  onMoveDown?: () => void;
  /** Whether this row can move up (e.g. not first) */
  canMoveUp?: boolean;
  /** Whether this row can move down (e.g. not last) */
  canMoveDown?: boolean;
  /** Card style overrides */
  style?: React.CSSProperties;
}

/**
 * Generic card row used for agent list and suggestion prompts list.
 * Same layout as agent card: optional order arrows (up/down) on the left, then image slot, header, action.
 */
export const ListCardRow: React.FunctionComponent<IListCardRowProps> = ({
  image,
  header,
  action,
  onMoveUp,
  onMoveDown,
  canMoveUp = true,
  canMoveDown = true,
  style
}) => {
  const showOrder = onMoveUp !== undefined || onMoveDown !== undefined;
  const leftContent = (
    <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
      {showOrder ? (
        <div
          style={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            gap: 0
          }}
        >
          {onMoveUp ? (
            <Button
              appearance="subtle"
              size="small"
              icon={<ChevronUp24Regular />}
              aria-label={strings.OrderUp}
              disabled={!canMoveUp}
              onClick={onMoveUp}
            />
          ) : null}
          {onMoveDown ? (
            <Button
              appearance="subtle"
              size="small"
              icon={<ChevronDown24Regular />}
              aria-label={strings.OrderDown}
              disabled={!canMoveDown}
              onClick={onMoveDown}
            />
          ) : null}
        </div>
      ) : null}
      {image ?? null}
    </div>
  );

  const safeHeader = header !== null && header !== undefined && header !== false ? header : undefined;
  const safeAction = action !== null && action !== undefined && action !== false ? action : undefined;

  return (
    <Card
      style={{
        borderRadius: tokens.borderRadiusXLarge,
        backgroundColor: tokens.colorNeutralBackground2,
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        ...style
      }}
    >
      <CardHeader image={leftContent} header={safeHeader} action={safeAction} />
    </Card>
  );
};
