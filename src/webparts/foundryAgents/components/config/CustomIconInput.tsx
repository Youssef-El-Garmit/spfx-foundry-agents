/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Input } from '@fluentui/react-components';
import strings from 'FoundryAgentsWebPartStrings';
import { getFluentIcon } from '../chat/shared/AgentIcon';
import styles from './CustomIconInput.module.scss';

const PREVIEW_SIZE = 20;

export interface ICustomIconInputProps {
  value: string;
  onChange: (value: string) => void;
  disabled?: boolean;
}

export const CustomIconInput: React.FunctionComponent<ICustomIconInputProps> = ({
  value,
  onChange,
  disabled
}) => {
  const displayValue = value ?? '';
  const [imgError, setImgError] = React.useState(false);
  const hasValidUrl = displayValue.trim().length > 0;
  const showCustomImage = hasValidUrl && !imgError;
  const DefaultIcon = getFluentIcon('Bot24Regular');

  React.useEffect(() => {
    setImgError(false);
  }, [displayValue]);

  return (
    <div className={styles.root}>
      <span className={styles.preview} aria-hidden>
        {showCustomImage ? (
          <img
            src={displayValue}
            alt=""
            className={styles.previewImg}
            width={PREVIEW_SIZE}
            height={PREVIEW_SIZE}
            onError={() => setImgError(true)}
          />
        ) : (
          <DefaultIcon style={{ width: PREVIEW_SIZE, height: PREVIEW_SIZE }} />
        )}
      </span>
      <Input
        id="custom-icon-url"
        type="url"
        value={displayValue}
        onChange={(_e, data) => onChange(data.value ?? '')}
        placeholder={strings.CustomIconUrlPlaceholder}
        disabled={disabled}
        className={styles.input}
      />
    </div>
  );
};
