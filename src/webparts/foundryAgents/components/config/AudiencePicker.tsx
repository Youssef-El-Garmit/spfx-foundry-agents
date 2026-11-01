/* eslint-disable @typescript-eslint/triple-slash-reference */
/// <reference path="../../FoundryAgentsWebPartStrings.d.ts" />
import * as React from 'react';
import { Input, Spinner } from '@fluentui/react-components';
import { DismissRegular, PeopleTeamRegular, BuildingRegular, PersonRegular } from '@fluentui/react-icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import strings from 'FoundryAgentsWebPartStrings';
import type { IAudienceGroup } from '../../models/IAgentConfig';
import { AudienceService, type IAudienceSearchResult } from '../../services/AudienceService';
import styles from './AudiencePicker.module.scss';

// ─────────────────────────────────────────────────────────────────────────────
// Props
// ─────────────────────────────────────────────────────────────────────────────

export interface IAudiencePickerProps {
  /** Currently selected audience groups */
  value: IAudienceGroup[];
  /** Called when the selection changes */
  onChange: (groups: IAudienceGroup[]) => void;
  /** SPFx web part context – needed to call Graph & SP REST */
  context: WebPartContext;
  /** Input placeholder */
  placeholder?: string;
  /** Parent scroll container or portal mount node (for positioning) */
  mountNode?: HTMLElement;
}

// ─────────────────────────────────────────────────────────────────────────────
// Component
// ─────────────────────────────────────────────────────────────────────────────

export const AudiencePicker: React.FunctionComponent<IAudiencePickerProps> = ({
  value,
  onChange,
  context,
  placeholder = strings.AudienceGroupsPlaceholder
}) => {
  const [query, setQuery] = React.useState('');
  const [results, setResults] = React.useState<IAudienceSearchResult[]>([]);
  const [isSearching, setIsSearching] = React.useState(false);
  const [dropdownOpen, setDropdownOpen] = React.useState(false);
  const debounceRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const wrapperRef = React.useRef<HTMLDivElement>(null);

  const serviceRef = React.useRef<AudienceService | null>(null);
  if (!serviceRef.current) serviceRef.current = new AudienceService(context);
  const service = serviceRef.current;

  // ── Search ──────────────────────────────────────────────────────────────

  React.useEffect(() => {
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (!query.trim()) {
      setResults([]);
      setDropdownOpen(false);
      return;
    }
    debounceRef.current = setTimeout(async () => {
      setIsSearching(true);
      setDropdownOpen(true);
      try {
        const found = await service.searchAll(query);
        // Filter already-selected entries
        const selectedIds = new Set(value.map((v) => `${v.type}:${v.id}`));
        setResults(found.filter((r) => !selectedIds.has(`${r.type}:${r.id}`)));
      } finally {
        setIsSearching(false);
      }
    }, 300);
    return () => {
      if (debounceRef.current) clearTimeout(debounceRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query]);

  // ── Close dropdown on outside click ────────────────────────────────────

  React.useEffect(() => {
    const handleClick = (e: MouseEvent): void => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) {
        setDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, []);

  // ── Handlers ────────────────────────────────────────────────────────────

  const handleSelect = (result: IAudienceSearchResult): void => {
    const group: IAudienceGroup = {
      id: result.id,
      displayName: result.displayName,
      type: result.type,
      mail: result.mail
    };
    onChange([...value, group]);
    setQuery('');
    setResults([]);
    setDropdownOpen(false);
  };

  const handleRemove = (group: IAudienceGroup): void => {
    onChange(value.filter((g) => !(g.type === group.type && g.id === group.id)));
  };

  // ── Render ───────────────────────────────────────────────────────────────

  return (
    <div className={styles.audiencePicker} ref={wrapperRef}>
      {/* Selected tags */}
      {value.length > 0 && (
        <div className={styles.tagList}>
          {value.map((g) => (
            <span key={`${g.type}:${g.id}`} className={styles.tag} title={g.mail ?? g.displayName}>
              {g.type === 'user' ? <PersonRegular fontSize={14} /> : g.type === 'graph' ? <PeopleTeamRegular fontSize={14} /> : <BuildingRegular fontSize={14} />}
              <span className={styles.tagLabel}>{g.displayName}</span>
              <span className={styles.tagTypeBadge}>
                {g.type === 'user' ? 'User' : g.type === 'graph' ? 'Entra' : 'SP'}
              </span>
              <span
                className={styles.tagRemove}
                role="button"
                tabIndex={0}
                aria-label={strings.RemoveAudienceAriaLabel.replace('{0}', g.displayName)}
                onClick={() => handleRemove(g)}
                onKeyDown={(e) => e.key === 'Enter' && handleRemove(g)}
              >
                <DismissRegular fontSize={12} />
              </span>
            </span>
          ))}
        </div>
      )}

      {/* Search input + dropdown */}
      <div className={styles.searchRow}>
        <Input
          className={styles.searchInput}
          value={query}
          onChange={(_e, data) => setQuery(data.value ?? '')}
          placeholder={placeholder}
          onFocus={() => query.trim() && setDropdownOpen(true)}
          contentAfter={isSearching ? <Spinner size="tiny" /> : undefined}
        />
        {dropdownOpen && (
          <div className={styles.dropdown}>
            {isSearching ? (
              <div className={styles.loadingDots}>{strings.AudienceSearching}</div>
            ) : results.length === 0 ? (
              <div className={styles.dropdownEmpty}>{strings.AudienceNoResults.replace('{0}', query)}</div>
            ) : (
              results.map((r) => (
                <div
                  key={`${r.type}:${r.id}`}
                  className={styles.dropdownItem}
                  role="option"
                  tabIndex={0}
                  onClick={() => handleSelect(r)}
                  onKeyDown={(e) => e.key === 'Enter' && handleSelect(r)}
                >
                  {r.type === 'user' ? <PersonRegular fontSize={16} /> : r.type === 'graph' ? <PeopleTeamRegular fontSize={16} /> : <BuildingRegular fontSize={16} />}
                  <span className={styles.dropdownItemName}>{r.displayName}</span>
                  {r.mail && (
                    <span className={styles.dropdownItemMeta}>{r.mail}</span>
                  )}
                  <span className={styles.dropdownItemMeta}>
                    {r.type === 'user' ? strings.AudienceTypeUser : r.type === 'graph' ? strings.AudienceTypeEntraGroup : strings.AudienceTypeSharePointGroup}
                  </span>
                </div>
              ))
            )}
          </div>
        )}
      </div>
    </div>
  );
};
