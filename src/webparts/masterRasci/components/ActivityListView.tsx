import * as React from 'react';
import styles from './MasterRasci.module.scss';
import ColumnResizer from './ColumnResizer';
import { IMasterRasciData, IRasciActivity } from './IMasterRasciProps';
import {
  CODES, HIGHLIGHT, GroupedRow, columnEntries, groupRows,
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
}

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

const ActivityRow: React.FC<{
  row: IRasciActivity;
  data: IMasterRasciData;
  role: number;
  collapseAfter: number;
  open: boolean;
  onToggle: (id: number) => void;
}> = ({ row, data, role, collapseAfter, open, onToggle }) => {
  const cols = CODES.map((c) => ({ meta: c, entries: columnEntries(row, c.code, role) }));
  // Only a row that is hiding something (or has been opened) has anything to toggle.
  const expandable = open || cols.some((c) => c.entries.length > collapseAfter);

  const toggle = (): void => onToggle(row.id);

  return (
    <div
      className={cx(styles.listGrid, styles.row, open && styles.rowOpen, expandable && styles.rowToggle)}
      // Clicking anywhere on the row expands every cell in it.
      onClick={expandable ? toggle : undefined}
      onKeyDown={expandable ? (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
      } : undefined}
      role={expandable ? 'button' : undefined}
      tabIndex={expandable ? 0 : undefined}
      aria-expanded={expandable ? open : undefined}
    >
      <div className={styles.rowTitleCell}>
        <div className={styles.rowTitle}>{row.activity || row.process}</div>
        {row.activity && <div className={styles.rowSub}>{row.process}</div>}
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
                    boxShadow: isMine ? `0 0 0 2px #0E1013, 0 0 0 4px ${HIGHLIGHT}` : 'none',
                    fontWeight: isMine ? 700 : 400,
                    opacity: role >= 0 && !isMine ? 0.5 : 1,
                  }}
                >
                  {data.roles[e.role].name}{e.code === 'AR' ? ' (A/R)' : ''}
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
            Process
            <ColumnResizer onChange={onActivityWidthChange} />
          </div>
          {CODES.map((c) => (
            <div key={c.code} className={styles.listHeadCode}>
              <span className={styles.badge} style={{ background: c.color, color: c.fg }}>{c.code}</span>
              {c.name}
            </div>
          ))}
        </div>

        {grouped.map((g) => g.kind === 'head' ? (
          <div key={g.key} className={styles.groupHead}>
            <span className={styles.groupHeadLabel}>
              {g.title}
              <span className={styles.groupCount}>{g.count} {g.count === 1 ? 'process' : 'processes'}</span>
            </span>
          </div>
        ) : (
          <ActivityRow
            key={g.key}
            row={g.row}
            data={data}
            role={role}
            collapseAfter={collapseAfter}
            open={!!open[g.row.id]}
            onToggle={toggle}
          />
        ))}
      </div>
    </div>
  );
};

export default ActivityListView;
