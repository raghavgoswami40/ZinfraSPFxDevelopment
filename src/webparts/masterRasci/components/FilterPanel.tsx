import * as React from 'react';
import styles from './MasterRasci.module.scss';
import { IFilterState, IMasterRasciData, IMasterRasciLabels } from './IMasterRasciProps';
import {
  EMPTY_FILTERS, hasActiveFilters, phaseOptions, groupOptions, phaseOfGroup,
} from './masterRasciModel';

interface IFilterPanelProps {
  data: IMasterRasciData;
  filters: IFilterState;
  onChange: (next: IFilterState) => void;
  labels: IMasterRasciLabels;
}

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

const FilterPanel: React.FC<IFilterPanelProps> = ({ data, filters, onChange, labels }) => {
  const phases = React.useMemo(() => phaseOptions(data), [data]);
  const groups = React.useMemo(() => groupOptions(data), [data]);

  const set = (patch: Partial<IFilterState>): void => onChange({ ...filters, ...patch });

  const onGroup = (value: string): void => {
    // Picking a group also selects its phase; "All process groups" leaves the
    // phase as it is.
    const phase = value === 'All' ? undefined : phaseOfGroup(data, value);
    set({ group: value, phase: phase !== undefined ? phase : filters.phase });
  };

  return (
    <div className={cx(styles.panel, styles.filters)} role="search" aria-label="Filter the Master RASCI">
      <div className={styles.phaseRow} role="group" aria-label={labels.phase}>
        <span className={styles.phaseLabel}>{labels.phase}</span>
        {phases.map((p) => {
          const on = filters.phase === p.key;
          return (
            <button
              key={p.key}
              type="button"
              aria-pressed={on}
              className={cx(styles.phaseBtn, on && styles.phaseBtnOn)}
              // A new phase starts from all of its process groups.
              onClick={() => set({ phase: p.key, group: 'All' })}
            >
              <span>{p.label}</span>
              <span className={styles.phaseCount}>{p.count}</span>
            </button>
          );
        })}
      </div>

      <div className={styles.divider} />

      <div className={styles.fieldRow}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>{labels.search}</span>
          <input
            type="text"
            className={styles.control}
            value={filters.q}
            placeholder="e.g. permit, risk, contract"
            onChange={(e) => set({ q: e.target.value })}
          />
        </label>

        <label className={styles.field}>
          <span className={styles.fieldLabel}>{labels.group}</span>
          <select
            className={cx(styles.control, styles.select)}
            value={filters.group}
            onChange={(e) => onGroup(e.target.value)}
          >
            <option value="All">All process groups</option>
            {groups.map((g) => (
              <optgroup key={g.phase} label={g.phase}>
                {g.groups.map((name) => <option key={name} value={name}>{name}</option>)}
              </optgroup>
            ))}
          </select>
        </label>

        <label className={styles.field}>
          <span className={styles.fieldLabel}>{labels.role}</span>
          <select
            className={cx(styles.control, styles.select)}
            value={filters.role < 0 ? '' : String(filters.role)}
            onChange={(e) => set({ role: e.target.value === '' ? -1 : Number(e.target.value) })}
          >
            <option value="">All roles</option>
            {data.roles.map((r, i) => (
              <option key={r.name} value={String(i)}>{r.name} ({r.group})</option>
            ))}
          </select>
        </label>

        {hasActiveFilters(filters) && (
          <button type="button" className={styles.clearBtn} onClick={() => onChange(EMPTY_FILTERS)}>
            Clear all
          </button>
        )}
      </div>
    </div>
  );
};

export default FilterPanel;
