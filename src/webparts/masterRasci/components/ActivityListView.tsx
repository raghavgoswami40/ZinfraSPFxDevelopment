import * as React from 'react';
import styles from './MasterRasci.module.scss';
import ColumnResizer from './ColumnResizer';
import { ChevronIcon, CollapseAllIcon, ExpandAllIcon } from './icons';
import { IMasterRasciData, IRasciActivity } from './IMasterRasciProps';
import {
  LIST_COLUMNS, HIGHLIGHT, GroupedRow, columnEntries, groupRows, isHidden, phaseKey,
} from './masterRasciModel';

interface IActivityListViewProps {
  data: IMasterRasciData;
  rows: IRasciActivity[];
  /** Selected role index, or -1. */
  role: number;
  /** Roles shown per cell before "+N more". */
  collapseAfter: number;
  /** Activity column width in px; undefined = the default flexible width. */
  activityWidth: number | undefined;
  onActivityWidthChange: (width: number | undefined) => void;
  /** Process groups the reader has folded up, by name. */
  collapsed: { [group: string]: boolean };
  onToggleGroup: (group: string) => void;
  onToggleAll: () => void;
  allCollapsed: boolean;
}

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

/** An L3 that has L4 steps beneath it and no RASCI of its own: a heading row with
 *  empty cells, so the steps below read as its group. */
const L3HeadRow: React.FC<{ title: string }> = ({ title }) => (
  <div className={cx(styles.listGrid, styles.row, styles.rowHeadOnly)}>
    <div className={styles.rowTitleCell}>
      <div className={cx(styles.rowTitle, styles.rowTitleL3)}>{title}</div>
    </div>
    {LIST_COLUMNS.map((c) => <div key={c.code} className={styles.codeCell} />)}
  </div>
);

const ActivityRow: React.FC<{
  row: IRasciActivity;
  data: IMasterRasciData;
  role: number;
  collapseAfter: number;
  open: boolean;
  onToggle: (id: number) => void;
  /** 3 = a process, 4 = a step beneath one. */
  level: 3 | 4;
  /** An L3 whose steps follow it. */
  parent: boolean;
  /** Last row of its L3 group: the border under it is the stronger one. */
  blockEnd: boolean;
}> = ({ row, data, role, collapseAfter, open, onToggle, level, parent, blockEnd }) => {
  const cols = LIST_COLUMNS.map((c) => ({ meta: c, entries: columnEntries(row, c.code, role) }));
  // Only a row that is hiding something (or has been opened) has anything to toggle.
  const expandable = open || cols.some((c) => c.entries.length > collapseAfter);

  const toggle = (): void => onToggle(row.id);

  return (
    <div
      className={cx(
        styles.listGrid, styles.row,
        parent && styles.rowParent,
        blockEnd && styles.rowEnd,
        open && styles.rowOpen,
        expandable && styles.rowToggle
      )}
      // Clicking anywhere on the row expands every cell in it.
      onClick={expandable ? toggle : undefined}
      onKeyDown={expandable ? (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
      } : undefined}
      role={expandable ? 'button' : undefined}
      tabIndex={expandable ? 0 : undefined}
      aria-expanded={expandable ? open : undefined}
    >
      <div className={cx(styles.rowTitleCell, level === 4 && styles.rowTitleCellL4)}>
        <div className={cx(styles.rowTitle, level === 4 ? styles.rowTitleL4 : styles.rowTitleL3)}>
          {level === 4 ? row.activity : row.process}
        </div>
        {open && <div className={styles.showLess} aria-hidden="true">Show less</div>}
      </div>

      {cols.map(({ meta, entries }) => {
        const shown = open ? entries : entries.slice(0, collapseAfter);
        const hidden = entries.length - shown.length;
        const holdsRole = entries.some((e) => e.role === role);
        return (
          <div key={meta.code} className={cx(styles.codeCell, holdsRole && styles.codeCellMine)}>
            {shown.map((e) => {
              const isMine = e.role === role;
              return (
                <span
                  key={e.role}
                  className={styles.chip}
                  style={{
                    background: isMine ? HIGHLIGHT : meta.chipBg,
                    color: isMine ? '#000' : meta.chipFg,
                    borderColor: isMine ? HIGHLIGHT : meta.color,
                    boxShadow: isMine ? `0 0 0 2px var(--mr-card), 0 0 0 4px ${HIGHLIGHT}` : 'none',
                    fontWeight: isMine ? 700 : 600,
                    opacity: role >= 0 && !isMine ? 0.5 : 1,
                  }}
                >
                  {data.roles[e.role].name}
                </span>
              );
            })}
            {hidden > 0 && <span className={styles.more}>+{hidden} more</span>}
            {entries.length === 0 && <span className={styles.dash}>&ndash;</span>}
          </div>
        );
      })}
    </div>
  );
};

const ActivityListView: React.FC<IActivityListViewProps> = ({
  data, rows, role, collapseAfter, activityWidth, onActivityWidthChange,
  collapsed, onToggleGroup, onToggleAll, allCollapsed,
}) => {
  const [open, setOpen] = React.useState<{ [id: number]: boolean }>({});
  const grouped: GroupedRow[] = React.useMemo(() => groupRows(rows), [rows]);

  const toggle = React.useCallback((id: number): void => {
    setOpen((prev) => ({ ...prev, [id]: !prev[id] }));
  }, []);

  return (
    <div className={styles.scrollX}>
      <div
        className={styles.listInner}
        style={activityWidth !== undefined ? ({ '--act-col': `${activityWidth}px` } as React.CSSProperties) : undefined}
      >
        <div className={cx(styles.listGrid, styles.listHeader)}>
          <div className={styles.listHeadActivity}>
            <span>Process</span>
            <button
              type="button"
              className={styles.collapseAllBtn}
              onClick={onToggleAll}
              title={allCollapsed ? 'Expand all' : 'Collapse all'}
              aria-label={allCollapsed ? 'Expand all' : 'Collapse all'}
            >
              {allCollapsed ? <ExpandAllIcon /> : <CollapseAllIcon />}
            </button>
            <ColumnResizer onChange={onActivityWidthChange} />
          </div>
          {LIST_COLUMNS.map((c) => (
            <div key={c.code} className={styles.listHeadCode}>
              <span className={styles.badge} style={{ background: c.color, color: c.fg }}>{c.code}</span>
              {c.name}
            </div>
          ))}
        </div>

        {grouped.map((g) => {
          if (g.kind === 'phase') {
            const key = phaseKey(g.title);
            return (
              <div key={g.key} className={cx(styles.listGrid, styles.phaseHead)} onClick={() => onToggleGroup(key)}>
                <div className={styles.phaseHeadLead}>
                  <button type="button" className={styles.phaseHeadLabel} aria-expanded={!collapsed[key]}>
                    <span className={cx(styles.chev, collapsed[key] && styles.chevCollapsed)}><ChevronIcon /></span>
                    {g.title}
                  </button>
                </div>
                <div className={styles.phaseHeadRest} />
              </div>
            );
          }
          if (g.kind === 'head') {
            if (collapsed[phaseKey(g.phase)]) { return null; }
            // The whole bar is the click target; the button inside is the keyboard stop
            // (its click bubbles up to the bar, so Enter and Space toggle once).
            return (
              <div key={g.key} className={cx(styles.listGrid, styles.groupHead)} onClick={() => onToggleGroup(g.title)}>
                <div className={styles.groupHeadLead}>
                  <button type="button" className={styles.groupHeadLabel} aria-expanded={!collapsed[g.title]}>
                    <span className={cx(styles.chev, collapsed[g.title] && styles.chevCollapsed)}><ChevronIcon /></span>
                    {g.title}
                  </button>
                </div>
                <div className={styles.groupHeadRest} />
              </div>
            );
          }
          if (g.kind === 'l3') {
            return isHidden(collapsed, g.phase, g.group) ? null : <L3HeadRow key={g.key} title={g.title} />;
          }
          if (isHidden(collapsed, g.row.phase, g.row.group)) { return null; }
          return (
            <ActivityRow
              key={g.key}
              row={g.row}
              data={data}
              role={role}
              collapseAfter={collapseAfter}
              open={!!open[g.row.id]}
              onToggle={toggle}
              level={g.level}
              parent={g.parent}
              blockEnd={g.blockEnd}
            />
          );
        })}
      </div>
    </div>
  );
};

export default ActivityListView;
