/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { memo } from 'react';
import { Tooltip } from '@fluentui/react-components';
import strings from 'FoundryAgentsWebPartStrings';
import type { IAnnotation } from '../../../types/chat';
import styles from './CitationMarker.module.scss';

interface CitationMarkerProps {
  /** 1-based citation index */
  index: number;
  /** The annotation data */
  annotation?: IAnnotation;
  /** Callback when marker is clicked */
  onClick: (index: number, annotation?: IAnnotation) => void;
}

function CitationMarkerComponent({ index, annotation, onClick }: CitationMarkerProps): React.ReactElement {
  const tooltipContent = annotation?.label || strings.CitationDefaultLabel.replace('{0}', String(index));

  const handleClick = (): void => {
    onClick(index, annotation);
  };

  const handleKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      handleClick();
    }
  };

  return (
    <Tooltip content={tooltipContent} relationship="label" withArrow>
      <sup
        className={styles.citationMarker}
        onClick={handleClick}
        onKeyDown={handleKeyDown}
        role="button"
        tabIndex={0}
        aria-label={strings.CitationAriaLabel.replace('{0}', String(index)).replace('{1}', tooltipContent)}
      >
        {index}
      </sup>
    </Tooltip>
  );
}

export const CitationMarker = memo(CitationMarkerComponent);
