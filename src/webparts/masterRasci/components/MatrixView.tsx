import * as React from 'react';
import styles from './MasterRasci.module.scss';
import { IMasterRasciData, IRasciActivity, RasciCode } from './IMasterRasciProps';
import ColumnResizer from './ColumnResizer';
import {
  CODES, HIGHLIGHT, groupRows, visibleRoles, roleBands,
} from './masterRasciModel';

interface IMatrixViewProps {
  data: IMasterRasciData;
  rows: IRasciActivity[];
  /** Selected role index, or -1. */
  role: number;
  /** Activity column width in px; undefined = the default width. */
  activityWidth: number | undefined;
  onActivityWidthChange: (width: number | undefined) => void;
}

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

/** Badge colours per code; AR takes Accountable's. */
const badgeColors = (code: RasciCode): { bg: string; fg: string } => {
  const meta = CODES[code === 'AR' ? 0 : ['A', 'R', 'S', 'C', 'I'].indexOf(code)];
  return { bg: meta.color, fg: meta.fg };
};

const MatrixView: React.FC<IMatrixViewProps> = ({
  data, rows, role, activityWidth, onActivityWidthChange,
}) => {
  // Only roles with an assignment somewhere in the current result get a column.
  const roles = React.useMemo(() => visibleRoles(data, rows), [data, rows]);
  const bands = React.useMemo(() => roleBands(data, roles), [data, roles]);
  const grouped = React.useMemo(() => groupRows(rows), [rows]);

  return (
    <div className={styles.matrixScroll}>
      <div
        className={styles.matrixGrid}
        style={{ gridTemplateColumns: `${activityWidth !== undefined ? `${activityWidth}px` : 'minmax(280px,340px)'} repeat(${roles.length},44px)` }}
      >
        <div className={styles.mCorner}>
          Process
          <ColumnResizer onChange={onActivityWidthChange} />
        </div>
        {bands.map((b) => (
          <div key={b.name} className={styles.mBand} style={{ gridColumn: `span ${b.span}` }}>{b.name}</div>
        ))}

        <div className={styles.mCornerSpacer} />
        {roles.map((i) => (
          <div key={i} className={cx(styles.mRoleHead, i === role && styles.mRoleHeadMine)}>
            <span className={styles.mRoleName}>{data.roles[i].name}</span>
          </div>
        ))}

        {grouped.map((g) => g.kind === 'head' ? (
          <div key={g.key} className={styles.mGroupHead}>
            <span className={styles.mGroupHeadLabel}>{g.title}</span>
          </div>
        ) : (
          <React.Fragment key={g.key}>
            <div className={styles.mActivity}>
              <div className={styles.mActivityTitle}>{g.row.activity || g.row.process}</div>
              {g.row.activity && <div className={styles.mActivitySub}>{g.row.process}</div>}
            </div>
            {roles.map((i) => {
              const code = g.row.assignments[i];
              const mine = i === role;
              const colors = code ? badgeColors(code) : undefined;
              return (
                <div key={i} className={cx(styles.mCell, mine && styles.mCellMine)}>
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
        ))}
      </div>
    </div>
  );
};

export default MatrixView;
