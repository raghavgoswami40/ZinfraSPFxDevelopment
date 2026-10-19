import * as React from 'react';
import styles from './MasterRasci.module.scss';
import { IMasterRasciData, IRasciActivity, RasciCode } from './IMasterRasciProps';
import ColumnResizer from './ColumnResizer';
import { scaleOf } from './scale';
import { ChevronIcon, CollapseAllIcon, ExpandAllIcon } from './icons';
import {
  CODES, HIGHLIGHT, groupRows, visibleRoles, roleBands, isHidden, phaseKey,
} from './masterRasciModel';

interface IMatrixViewProps {
  data: IMasterRasciData;
  rows: IRasciActivity[];
  /** Selected role index, or -1. */
  role: number;
  /** Activity column width in px; undefined = the default width. */
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

/** Badge colours per code; AR takes Accountable's. */
const badgeColors = (code: RasciCode): { bg: string; fg: string } => {
  const meta = CODES[code === 'AR' ? 0 : ['A', 'R', 'S', 'C', 'I'].indexOf(code)];
  return { bg: meta.color, fg: meta.fg };
};

/** "Project Leadership" -> "PL". */
const initials = (name: string): string =>
  name.split(/\s+/).filter(Boolean).map((w) => w.charAt(0).toUpperCase()).join('').slice(0, 3);

const MatrixView: React.FC<IMatrixViewProps> = ({
  data, rows, role, activityWidth, onActivityWidthChange,
  collapsed, onToggleGroup, onToggleAll, allCollapsed,
}) => {
  // Only roles with an assignment somewhere in the current result get a column.
  const roles = React.useMemo(() => visibleRoles(data, rows), [data, rows]);
  const bands = React.useMemo(() => roleBands(data, roles), [data, roles]);
  const grouped = React.useMemo(() => groupRows(rows), [rows]);

  // Role names read horizontally (wrapping onto up to three lines) when every column has room
  // for them, and vertically otherwise, so a narrower screen never gets unreadable headers.
  // "Room" is what is left after the Process column, shared between the roles. Once horizontal,
  // it stays horizontal until the room drops a little below the threshold, so a scrollbar
  // appearing or disappearing at the boundary cannot make the headers flicker between the two.
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const [horizontal, setHorizontal] = React.useState(false);
  React.useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el) { return undefined; }
    const measure = (): void => {
      const k = scaleOf(el);
      const processMin = activityWidth !== undefined ? activityWidth : 280 * k;
      const room = (el.clientWidth - processMin) / Math.max(1, roles.length);
      setHorizontal((was) => (was ? room >= 80 * k : room >= 88 * k));
    };
    measure();
    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', measure);
      return () => window.removeEventListener('resize', measure);
    }
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [roles.length, activityWidth]);

  // Fit the available width. Role columns grow from 44 to at most 120 scale units and the
  // Process column from 280 to at most 560, so a matrix with many roles fills its box on a wide
  // screen instead of leaving an empty strip on the right. The caps keep a matrix with only a
  // few roles compact: otherwise the Process column would stretch across the screen and put
  // names far from their badges. The grid never gets narrower than the minimums, so with many
  // roles on a narrow screen it scrolls sideways as before. A Process width the reader has
  // dragged is kept as it is.
  const U = 'var(--u)';
  const minRole = `calc(${horizontal ? 88 : 44} * ${U})`;
  const maxRole = `calc(${horizontal ? 150 : 120} * ${U})`;
  const processTrack = activityWidth !== undefined ? `${activityWidth}px` : `minmax(calc(280 * ${U}), calc(560 * ${U}))`;
  const processMin = activityWidth !== undefined ? `${activityWidth}px` : `calc(280 * ${U})`;
  const roleTrack = `minmax(${minRole}, ${maxRole})`;
  const gridStyle: React.CSSProperties = {
    gridTemplateColumns: `${processTrack} repeat(${roles.length}, ${roleTrack})`,
    minWidth: `calc(${processMin} + ${roles.length} * ${minRole})`,
  };

  return (
    <div className={styles.matrixScroll} ref={scrollRef}>
      <div
        className={styles.matrixGrid}
        style={gridStyle}
      >
        <div className={styles.mCorner}>
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
        {bands.map((b) => (
          <div
            key={b.name}
            className={styles.mBand}
            style={{ gridColumn: `span ${b.span}` }}
            // A one-column band is too narrow for a multi-word name, so that shows initials.
            title={b.name}
            aria-label={b.name}
          >
            {b.span === 1 && b.name.indexOf(' ') !== -1 ? initials(b.name) : b.name}
          </div>
        ))}

        <div className={cx(styles.mCornerSpacer, horizontal && styles.mCornerSpacerH)} />
        {roles.map((i) => (
          <div
            key={i}
            className={cx(styles.mRoleHead, horizontal && styles.mRoleHeadH, i === role && styles.mRoleHeadMine)}
          >
            <span className={horizontal ? styles.mRoleNameH : styles.mRoleName}>{data.roles[i].name}</span>
          </div>
        ))}

        {grouped.map((g) => {
          if (g.kind === 'phase') {
            const key = phaseKey(g.title);
            return (
              <React.Fragment key={g.key}>
                <div className={styles.mPhaseLead} onClick={() => onToggleGroup(key)}>
                  <button type="button" className={styles.phaseHeadLabel} aria-expanded={!collapsed[key]}>
                    <span className={cx(styles.chev, collapsed[key] && styles.chevCollapsed)}><ChevronIcon /></span>
                    {g.title}
                  </button>
                </div>
                <div className={styles.mPhaseRest} onClick={() => onToggleGroup(key)} />
              </React.Fragment>
            );
          }
          if (g.kind === 'head') {
            if (collapsed[phaseKey(g.phase)]) { return null; }
            return (
              <React.Fragment key={g.key}>
                <div className={styles.mGroupLead} onClick={() => onToggleGroup(g.title)}>
                  <button type="button" className={styles.mGroupHeadLabel} aria-expanded={!collapsed[g.title]}>
                    <span className={cx(styles.chev, collapsed[g.title] && styles.chevCollapsed)}><ChevronIcon /></span>
                    {g.title}
                  </button>
                </div>
                <div className={styles.mGroupRest} onClick={() => onToggleGroup(g.title)} />
              </React.Fragment>
            );
          }
          if (g.kind === 'l3') {
            // An L3 with steps and no RASCI of its own: a heading row, empty cells.
            return isHidden(collapsed, g.phase, g.group) ? null : (
              <React.Fragment key={g.key}>
                <div className={cx(styles.mActivity, styles.mActivityHead)}>
                  <div className={cx(styles.mActivityTitle, styles.mTitleL3)}>{g.title}</div>
                </div>
                {roles.map((i) => <div key={i} className={cx(styles.mCell, styles.mCellHead)} />)}
              </React.Fragment>
            );
          }
          if (isHidden(collapsed, g.row.phase, g.row.group)) { return null; }
          const isL4 = g.level === 4;
          const end = g.blockEnd ? styles.mEnd : undefined;
          return (
          <React.Fragment key={g.key}>
            <div className={cx(styles.mActivity, isL4 && styles.mActivityL4, g.parent && styles.mActivityHead, end)}>
              <div className={cx(styles.mActivityTitle, isL4 ? styles.mTitleL4 : styles.mTitleL3)}>
                {isL4 ? g.row.activity : g.row.process}
              </div>
            </div>
            {roles.map((i) => {
              const code = g.row.assignments[i];
              const mine = i === role;
              const colors = code ? badgeColors(code) : undefined;
              return (
                <div key={i} className={cx(styles.mCell, mine && styles.mCellMine, g.parent && styles.mCellHead, end)}>
                  {code && colors && (
                    <span
                      className={styles.mBadge}
                      style={{
                        background: colors.bg,
                        color: colors.fg,
                        opacity: role >= 0 && !mine ? 0.4 : 1,
                        boxShadow: mine ? `0 0 0 2px ${HIGHLIGHT}` : 'none',
                      }}
                    >
                      {code === 'AR' ? 'A/R' : code}
                    </span>
                  )}
                </div>
              );
            })}
          </React.Fragment>
          );
        })}
      </div>
    </div>
  );
};

export default MatrixView;
