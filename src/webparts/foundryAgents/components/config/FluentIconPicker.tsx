/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Combobox, Option, useId } from '@fluentui/react-components';
import strings from 'FoundryAgentsWebPartStrings';
import { getFluentIcon, ICON_NAMES } from '../chat/shared/AgentIcon';
import styles from './FluentIconPicker.module.scss';

export interface IFluentIconPickerProps {
  value: string;
  onChange: (value: string) => void;
  disabled?: boolean;
}

const MAX_VISIBLE = 100;

export const FluentIconPicker: React.FunctionComponent<IFluentIconPickerProps> = ({
  value,
  onChange,
  disabled
}) => {
  const id = useId('fluent-icon-picker');
  const [inputValue, setInputValue] = React.useState(value || '');

  React.useEffect(() => {
    setInputValue(value || '');
  }, [value]);

  const filteredNames = React.useMemo(() => {
    const q = inputValue.toLowerCase();
    return q
      ? ICON_NAMES.filter((n) => n.toLowerCase().includes(q)).slice(0, MAX_VISIBLE)
      : ICON_NAMES.slice(0, MAX_VISIBLE);
  }, [inputValue]);

  const SelectedIcon = getFluentIcon(value || 'Bot24Regular');

  return (
    <div className={styles.root}>
      <div className={styles.comboboxWrapper}>
        <span className={styles.preview} aria-hidden>
          <SelectedIcon style={{ width: 20, height: 20 }} />
        </span>
        <Combobox
          id={id}
          freeform
          inlinePopup
          value={inputValue}
          selectedOptions={value ? [value] : []}
          onChange={(e) => setInputValue(e.target.value)}
          onOptionSelect={(_e, data) => {
            const v = data.optionValue ?? '';
            setInputValue(v);
            onChange(v);
          }}
          disabled={disabled}
          placeholder={strings.FluentIconPlaceholder}
          className={styles.combobox}
        >
          {filteredNames.map((name) => {
            const Icon = getFluentIcon(name);
            return (
              <Option key={name} value={name} text={name}>
                <span className={styles.optionContent}>
                  <Icon style={{ width: 16, height: 16, flexShrink: 0 }} />
                  <span>{name}</span>
                </span>
              </Option>
            );
          })}
        </Combobox>
      </div>
    </div>
  );
};
