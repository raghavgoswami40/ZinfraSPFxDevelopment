import * as React from 'react';
import styles from './Rasci.module.scss';
import {
  IRasciProps, RoleKey, ROLE_ORDER, ROLE_META,
} from './IRasciProps';
import { textOn, onDark } from './roleColors';

const DEFAULT_TITLE = 'RASCI';
const DEFAULT_ICON  = '👥';
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

/** One role: its heading and the pills beneath it. */
const RoleSection: React.FC<{
  role: RoleKey;
  holders: string[];
  color: string;
}> = ({ role, holders, color }) => {
  const label = ROLE_META[role].label;
  const tagText = textOn(color);

  return (
    <div className={styles.roleBlock}>
      <div className={styles.roleLabel} style={{ color: onDark(color) }}>
        <span className={styles.roleSwatch} style={{ background: color }} aria-hidden="true" />
        {label}
      </div>
      {holders.length === 0 ? (
        <div className={styles.roleEmpty}>N/A</div>
      ) : (
        <div className={styles.tags}>
          {holders.map((holder, i) => (
            <span
              key={`${holder}_${i}`}
              className={styles.tag}
              style={{ background: color, color: tagText }}
            >
              {holder}
            </span>
          ))}
        </div>
      )}
    </div>
  );
};

const Rasci: React.FC<IRasciProps> = (props) => {
  const {
    record,
    isLoading,
    error,
    emptyReason,
    notFoundCode,
    onRetry,
    title = DEFAULT_TITLE,
    headerIcon = DEFAULT_ICON,
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
        <div className={styles.titleBlock}>
          <h3 className={styles.stepTitle}>{record.title}</h3>
        </div>

        {ROLE_ORDER.map((role) => (
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
    <div
      className={styles.panel}
      role="region"
      aria-label={title || DEFAULT_TITLE}
      data-sticky-role="rasci"
    >
      <div className={styles.header}>
        {headerIcon && <span className={styles.headerIcon} aria-hidden="true">{headerIcon}</span>}
        <span className={styles.headerTitle}>{title || DEFAULT_TITLE}</span>
      </div>
      <div className={styles.body}>
        {renderBody()}
      </div>
    </div>
  );
};

export default Rasci;
