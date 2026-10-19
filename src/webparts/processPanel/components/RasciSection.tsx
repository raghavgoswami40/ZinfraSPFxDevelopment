import * as React from 'react';
import styles from './RasciSection.module.scss';
import {
  IRasciSectionProps, RoleKey, ROLE_ORDER, ROLE_META,
} from './IProcessPanelProps';
import { onDark } from './roleColors';
import { RasciLogo } from './SectionIcons';

const DEFAULT_TITLE = 'RASCI';
const DEFAULT_EMPTY_ICON = '👈';
const DEFAULT_EMPTY_ICON_L4 = '👆';

const DEFAULT_EMPTY_HEADING_L3 = 'Select a process step to view its RASCI';
const DEFAULT_EMPTY_BODY_L3 =
  'Click on any process in the process hierarchy to see who is responsible, ' +
  'accountable, supporting, consulted and informed.';
const DEFAULT_EMPTY_HEADING_L4 = 'Select an L4 process step to view its RASCI';
const DEFAULT_EMPTY_BODY_L4 =
  'This process has sub-processes. Click one of the L4 process names in the ' +
  'process details panel to see its RASCI.';

/** Black or white, whichever reads better on a role-coloured fill. */
const textOn = (fill: string): string => {
  const m = /^#?([0-9a-f]{6}|[0-9a-f]{3})$/i.exec((fill || '').trim());
  if (!m) { return '#ffffff'; }
  const h = m[1].length === 3 ? m[1].split('').map((c) => c + c).join('') : m[1];
  const lum = (0.299 * parseInt(h.slice(0, 2), 16) + 0.587 * parseInt(h.slice(2, 4), 16) +
    0.114 * parseInt(h.slice(4, 6), 16)) / 255;
  return lum > 0.6 ? '#1c1c1c' : '#ffffff';
};

/** Popup layout: one card per role, with a letter tile and a chip per person. */
const RoleCard: React.FC<{
  role: RoleKey;
  holders: string[];
  color: string;
}> = ({ role, holders, color }) => {
  const label = ROLE_META[role].label;
  const empty = holders.length === 0;
  const cls = [
    styles.roleCard,
    empty ? styles.roleCardEmpty : '',
  ].filter(Boolean).join(' ');

  return (
    <div className={cls} style={empty ? undefined : { borderTopColor: color }}>
      <div className={styles.roleCardHead}>
        <span
          className={styles.roleTile}
          style={empty ? undefined : { background: color, color: textOn(color) }}
          aria-hidden="true"
        >
          {label.charAt(0)}
        </span>
        <span className={styles.roleCardName} style={empty ? undefined : { color: onDark(color) }}>
          {label}
        </span>
      </div>
      {empty && <div className={styles.roleCardNone}>N/A</div>}
      {!empty && (
        <ul className={styles.chips}>
          {holders.map((holder, i) => (
            <li key={`${holder}_${i}`} className={styles.chip}>
              {holder}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};

/** One role: its heading and the pills beneath it. */
const RoleSection: React.FC<{
  role: RoleKey;
  holders: string[];
  color: string;
}> = ({ role, holders, color }) => {
  const label = ROLE_META[role].label;

  return (
    <div className={styles.roleBlock}>
      <div className={styles.roleLabel} style={{ color: onDark(color) }}>
        {label}
      </div>
      {holders.length === 0 ? (
        <div className={styles.roleEmpty}>N/A</div>
      ) : (
        <ul className={styles.holders}>
          {holders.map((holder, i) => <li key={`${holder}_${i}`}>{holder}</li>)}
        </ul>
      )}
    </div>
  );
};

const RasciSection: React.FC<IRasciSectionProps> = (props) => {
  const {
    compact,
    record,
    isLoading,
    error,
    emptyReason,
    notFoundCode,
    onRetry,
    title = DEFAULT_TITLE,
    colors,
    emptyIcon = DEFAULT_EMPTY_ICON,
    emptyIconL4 = DEFAULT_EMPTY_ICON_L4,
    emptyHeadingL3 = DEFAULT_EMPTY_HEADING_L3,
    emptyBodyL3 = DEFAULT_EMPTY_BODY_L3,
    emptyHeadingL4 = DEFAULT_EMPTY_HEADING_L4,
    emptyBodyL4 = DEFAULT_EMPTY_BODY_L4,
  } = props;

  const renderBody = (): JSX.Element => {
    if (isLoading) {
      return (
        <div className={styles.loading} aria-busy="true" aria-label="Loading RASCI">
          <div className={`${styles.shimmer} ${styles.shimmerA}`} />
          <div className={`${styles.shimmer} ${styles.shimmerB}`} />
          <div className={`${styles.shimmer} ${styles.shimmerC}`} />
        </div>
      );
    }

    if (error) {
      return (
        <div className={styles.empty}>
          <div className={styles.emptyHeading}>Couldn&rsquo;t load the RASCI</div>
          <div className={styles.emptyBody}>{error}</div>
          {onRetry && (
            <button type="button" className={styles.retryBtn} onClick={onRetry}>Retry</button>
          )}
        </div>
      );
    }

    if (!record) {
      // Three distinct "nothing to show" states. Collapsing them would leave a
      // reader who has just selected an L3 with sub-processes staring at
      // "select a process", having done exactly that.
      if (emptyReason === 'notFound') {
        return (
          <div className={styles.empty}>
            <div className={styles.emptyIcon} aria-hidden="true">📄</div>
            <div className={styles.emptyHeading}>No RASCI recorded for {notFoundCode}</div>
            <div className={styles.emptyBody}>
              This process has no row in the RASCI list yet.
            </div>
          </div>
        );
      }

      const awaitingL4 = emptyReason === 'awaitingL4';
      return (
        <div className={styles.empty}>
          <div className={styles.emptyIcon} aria-hidden="true">
            {awaitingL4 ? emptyIconL4 : emptyIcon}
          </div>
          <div className={styles.emptyHeading}>
            {awaitingL4 ? emptyHeadingL4 : emptyHeadingL3}
          </div>
          <div className={styles.emptyBody}>
            {awaitingL4 ? emptyBodyL4 : emptyBodyL3}
          </div>
        </div>
      );
    }

    return (
      <>
        {!compact && (
          <div className={styles.titleBlock}>
            <h3 className={styles.stepTitle}>
              {record.code ? `${record.code} ${record.title}` : record.title}
            </h3>
          </div>
        )}

        {compact ? (
          <div className={styles.cardGrid}>
            {ROLE_ORDER.map((role) => (
              <RoleCard key={role} role={role} holders={record.roles[role]} color={colors[role]} />
            ))}
          </div>
        ) : ROLE_ORDER.map((role) => (
          <RoleSection
            key={role}
            role={role}
            holders={record.roles[role]}
            color={colors[role]}
          />
        ))}
      </>
    );
  };

  return (
    <div className={styles.panel} role="region" aria-label={title || DEFAULT_TITLE}>
      {!compact && (
        <div className={styles.header}>
          <span className={styles.headerIcon} aria-hidden="true"><RasciLogo /></span>
          <span className={styles.headerTitle}>{title || DEFAULT_TITLE}</span>
        </div>
      )}
      <div className={styles.body}>
        {renderBody()}
      </div>
    </div>
  );
};

export default RasciSection;
