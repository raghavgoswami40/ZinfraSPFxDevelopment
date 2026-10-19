import * as React from 'react';
import styles from './MasterRasci.module.scss';
import { IFilterState, IMasterRasciData, IMasterRasciLabels, ChipCode } from './IMasterRasciProps';
import {
  EMPTY_FILTERS, LIST_COLUMNS, hasActiveFilters, phaseOptions, groupOptions, phaseOfGroup,
} from './masterRasciModel';
import GroupedPicker, { IPickerGroup } from './GroupedPicker';
import ThemeToggle from './ThemeToggle';
import { SearchIcon, CloseIcon, LayersIcon, UserIcon, FunnelIcon, ChevronIcon } from './icons';

interface IFilterPanelProps {
  data: IMasterRasciData;
  filters: IFilterState;
  onChange: (next: IFilterState) => void;
  labels: IMasterRasciLabels;
}

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

/** One removable "this filter is on" chip. */
const ActiveChip: React.FC<{ name: string; value: string; onRemove: () => void }> = ({
  name, value, onRemove,
}) => (
  <button
    type="button"
    className={styles.filterChip}
    onClick={onRemove}
    aria-label={`Remove filter ${name}: ${value}`}
  >
    <span className={styles.chipKey}>{name}</span>
    <span>{value}</span>
    <CloseIcon />
  </button>
);

let pickerSeq = 0;

const OPEN_KEY = 'mr-filters-open';
const readOpen = (): boolean => {
  try { return window.localStorage.getItem(OPEN_KEY) !== '0'; } catch { return true; }
};

const FilterPanel: React.FC<IFilterPanelProps> = ({ data, filters, onChange, labels }) => {
  // One id per instance, so two Master RASCI web parts on a page do not share labels.
  const seq = React.useRef(++pickerSeq).current;
  const roleLabelId = `mr-role-label-${seq}`;
  const groupLabelId = `mr-group-label-${seq}`;
  const codeLabelId = `mr-code-label-${seq}`;
  const phases = React.useMemo(() => phaseOptions(data), [data]);
  // Process groups under their phase, role names under their role group.
  const groupChoices = React.useMemo<IPickerGroup[]>(
    () => groupOptions(data).map((g) => ({
      name: g.phase,
      options: g.groups.map((name) => ({ value: name, label: name })),
    })),
    [data]
  );
  const roleChoices = React.useMemo<IPickerGroup[]>(() => {
    const out: IPickerGroup[] = [];
    data.roles.forEach((r, i) => {
      let g = out.length > 0 ? out[out.length - 1] : undefined;
      if (!g || g.name !== r.group) { g = { name: r.group, options: [] }; out.push(g); }
      g.options.push({ value: String(i), label: r.name });
    });
    return out;
  }, [data]);

  // The filter boxes fold away to give the table the room; the chips below stay, so a folded
  // panel still says what is being filtered.
  const [open, setOpen] = React.useState<boolean>(readOpen);
  const toggleOpen = (): void => {
    const next = !open;
    setOpen(next);
    try { window.localStorage.setItem(OPEN_KEY, next ? '1' : '0'); } catch { /* storage blocked */ }
  };
  const fieldsId = `mr-filter-fields-${seq}`;
  const activeCount = (filters.phase !== 'All' ? 1 : 0) + (filters.group !== 'All' ? 1 : 0)
    + (filters.role >= 0 ? 1 : 0) + (filters.code ? 1 : 0) + (filters.q ? 1 : 0);

  const set = (patch: Partial<IFilterState>): void => onChange({ ...filters, ...patch });

  const onGroup = (value: string): void => {
    // Picking a group also selects its phase; "All process groups" leaves the
    // phase as it is.
    const phase = value === 'All' ? undefined : phaseOfGroup(data, value);
    set({ group: value, phase: phase !== undefined ? phase : filters.phase });
  };

  // The five responsibility types, in the By process column order, each with its colour badge.
  const codeChoices = React.useMemo<IPickerGroup[]>(() => [{
    name: '',
    options: LIST_COLUMNS.map((c) => ({
      value: c.code,
      label: c.name,
      badge: { text: c.code, bg: c.color, fg: c.fg },
    })),
  }], []);

  const codeName = filters.code ? LIST_COLUMNS.filter((c) => c.code === filters.code)[0].name : '';
  const roleName = filters.role >= 0 && data.roles[filters.role] ? data.roles[filters.role].name : '';

  return (
    <div className={styles.toolbar} role="search" aria-label="Filter the Master RASCI">
      <div className={styles.tabsRow} role="group" aria-label={labels.phase}>
        <div className={styles.tabs}>
          {phases.map((p) => {
            const on = filters.phase === p.key;
            return (
              <button
                key={p.key}
                type="button"
                aria-pressed={on}
                className={cx(styles.tab, on && styles.tabOn)}
                // A new phase starts from all of its process groups.
                onClick={() => set({ phase: p.key, group: 'All' })}
              >
                <span>{p.label}</span>
                <span className={styles.tabCount}>{p.count}</span>
              </button>
            );
          })}
        </div>
        <button
          type="button"
          className={cx(styles.filtersToggle, open && styles.filtersToggleOpen)}
          aria-expanded={open}
          aria-controls={fieldsId}
          onClick={toggleOpen}
        >
          <FunnelIcon />
          <span>Filters</span>
          {activeCount > 0 && (
            <span className={styles.filtersCount} aria-label={`${activeCount} active`}>{activeCount}</span>
          )}
          <span className={styles.filtersChevron} aria-hidden="true"><ChevronIcon /></span>
        </button>
        <ThemeToggle />
      </div>

      {open && <div className={styles.fieldRow} id={fieldsId}>
        <label className={styles.field}>
          <span className={styles.fieldLabel}>{labels.search}</span>
          <span className={styles.searchWrap}>
            <span className={styles.searchIcon}><SearchIcon /></span>
            <input
              type="text"
              className={cx(styles.control, styles.searchInput)}
              value={filters.q}
              placeholder="e.g. permit, risk, contract"
              onChange={(e) => set({ q: e.target.value })}
            />
          </span>
        </label>

        <div className={styles.field}>
          <span className={styles.fieldLabel} id={groupLabelId}>{labels.group}</span>
          <GroupedPicker
            groups={groupChoices}
            value={filters.group}
            allValue="All"
            allLabel="All process groups"
            icon={<LayersIcon />}
            labelId={groupLabelId}
            listLabel="Process group"
            maxWidth={560}
            tone="accent"
            onChange={onGroup}
          />
        </div>

        <div className={styles.field}>
          <span className={styles.fieldLabel} id={roleLabelId}>{labels.role}</span>
          <GroupedPicker
            groups={roleChoices}
            value={filters.role < 0 ? '' : String(filters.role)}
            allValue=""
            allLabel="All roles"
            icon={<UserIcon />}
            labelId={roleLabelId}
            listLabel="Role"
            maxWidth={760}
            tone="highlight"
            onChange={(v) => set({ role: v === '' ? -1 : Number(v) })}
          />
        </div>

        <div className={styles.field}>
          <span className={styles.fieldLabel} id={codeLabelId}>{labels.code}</span>
          <GroupedPicker
            groups={codeChoices}
            value={filters.code}
            allValue=""
            allLabel="Any responsibility"
            icon={<FunnelIcon />}
            labelId={codeLabelId}
            listLabel="Responsibility"
            maxWidth={340}
            tone="accent"
            onChange={(v) => set({ code: v as ChipCode | '' })}
          />
        </div>
      </div>}

      {hasActiveFilters(filters) && (
        <div className={styles.activeRow} aria-label="Active filters">
          {filters.phase !== 'All' && (
            <ActiveChip name={labels.phase} value={filters.phase}
              onRemove={() => set({ phase: 'All', group: 'All' })} />
          )}
          {filters.group !== 'All' && (
            <ActiveChip name={labels.group} value={filters.group} onRemove={() => set({ group: 'All' })} />
          )}
          {roleName && (
            <ActiveChip name={labels.role} value={roleName} onRemove={() => set({ role: -1 })} />
          )}
          {filters.code && (
            <ActiveChip name={labels.code} value={codeName} onRemove={() => set({ code: '' })} />
          )}
          {filters.q && (
            <ActiveChip name={labels.search} value={`“${filters.q}”`} onRemove={() => set({ q: '' })} />
          )}
          <button type="button" className={styles.clearBtn} onClick={() => onChange(EMPTY_FILTERS)}>
            Clear all
          </button>
        </div>
      )}
    </div>
  );
};

export default FilterPanel;
