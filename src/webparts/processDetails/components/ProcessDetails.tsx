import * as React from 'react';
import styles from './ProcessDetails.module.scss';
import {
  IProcessDetailsProps, IOutputItem, ISubStep, IDiagnosticsView,
} from './IProcessDetailsProps';

const DEFAULT_TITLE = 'Process details';
const DEFAULT_ICON  = '🧾';
const DEFAULT_EMPTY_ICON    = '👈';
const DEFAULT_EMPTY_HEADING = 'Select a process to view details';
const DEFAULT_EMPTY_BODY    =
  'Click on any process in the process hierarchy to see its inputs, procedure and outputs here.';

// Ids must be unique per web part instance (two panels can share a page) and
// safe for an HTML id, so codes like "2.1.3.1" lose their dots.
const slug = (code: string): string => code.replace(/[^a-zA-Z0-9]+/g, '-');

/**
 * An output's artefact icon. Three cases:
 *  - logo + link  -> a real link, opens the document or SAP transaction
 *  - logo, no link -> shown but dimmed, so a missing URL is visible rather than
 *                     indistinguishable from an untagged output
 *  - no logo       -> nothing (the system had no System Master match)
 *
 * The logo is an <img> at a fixed height with automatic width, not a
 * fixed-size background-image: logos come from a SharePoint list with arbitrary
 * aspect ratios, and a fixed-width box would letterbox the wide ones.
 */
const OutputIcon: React.FC<{ item: IOutputItem }> = ({ item }) => {
  if (!item.logoUrl) { return null; }

  const img = <img className={styles.outputLogo} src={item.logoUrl} alt="" />;

  if (!item.href) {
    return (
      <span
        className={`${styles.outputIcon} ${styles.outputIconDisabled}`}
        title={item.system}
      >
        {img}
      </span>
    );
  }

  const label = `Open ${item.text}${item.system ? ` in ${item.system}` : ''}`;
  return (
    <a
      className={styles.outputIcon}
      href={item.href}
      target="_blank"
      rel="noopener noreferrer"
      title={label}
      aria-label={label}
    >
      {img}
    </a>
  );
};

const OutputList: React.FC<{ items: IOutputItem[] }> = ({ items }) => (
  <ul className={styles.list}>
    {items.map((item, i) => (
      <li key={`${item.text}_${i}`}>
        {item.text}
        <OutputIcon item={item} />
      </li>
    ))}
  </ul>
);

const Diagnostics: React.FC<{ data: IDiagnosticsView }> = ({ data }) => {
  const c = data.connection;
  // The two states that account for almost every "the panel shows nothing":
  // never connected, or connected but no click has arrived yet.
  const notConnected = !c.sourceTitle;

  return (
  <div className={styles.diagnostics}>
    <div className={styles.diagHeading}>Diagnostics</div>

    {notConnected && (
      <div className={styles.diagWarn}>
        Not connected to a process hierarchy. Open the property pane &rarr;
        Connection, and pick a source under &ldquo;Connect to a process hierarchy&rdquo;.
      </div>
    )}
    {!notConnected && !c.hasValue && (
      <div className={styles.diagWarn}>
        Connected to &ldquo;{c.sourceTitle}&rdquo;, but no step has been selected yet.
        Switch the page out of edit mode, then click a rectangle.
      </div>
    )}

    <div className={styles.diagBlock}>
      <div className={styles.diagSubheading}>Connection</div>
      <table className={styles.diagTable}>
        <tbody>
          <tr><td>Source</td><td>{c.sourceTitle || <em>none</em>}</td></tr>
          <tr><td>Payload received</td><td>{c.hasValue ? 'yes' : 'no'}</td></tr>
          {c.hasValue && <tr><td>Payload code</td><td><code>{c.payloadCode || '(none)'}</code></td></tr>}
          {c.hasValue && <tr><td>Payload label</td><td>{c.payloadLabel}</td></tr>}
          <tr><td>Code in use</td><td><code>{c.effectiveCode || '(none)'}</code></td></tr>
          <tr><td>Fetch</td><td>{c.fetchState}</td></tr>
        </tbody>
      </table>
    </div>

    {data.unresolvedSystems.length > 0 && (
      <div className={styles.diagWarn}>
        No logo resolved for: {data.unresolvedSystems.join(', ')}
      </div>
    )}

    {[
      { label: data.processList, fields: data.processFields },
      { label: data.systemList, fields: data.systemFields },
    ].map((list) => (
      <div key={list.label} className={styles.diagBlock}>
        <div className={styles.diagSubheading}>{list.label} — columns</div>
        {list.fields.length === 0
          ? <div className={styles.diagEmpty}>not loaded</div>
          : (
            <table className={styles.diagTable}>
              <tbody>
                {list.fields.map((f) => (
                  <tr key={f.internalName}>
                    <td>{f.displayName}</td>
                    <td><code>{f.internalName}</code></td>
                    <td>{f.typeAsString}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
      </div>
    ))}

    <div className={styles.diagBlock}>
      <div className={styles.diagSubheading}>Systems with a logo</div>
      {data.systemLogos.length === 0
        ? <div className={styles.diagEmpty}>none</div>
        : (
          <ul className={styles.diagList}>
            {data.systemLogos.map((s) => (
              <li key={s.system}>
                <img className={styles.outputLogo} src={s.logoUrl} alt="" /> {s.system}
              </li>
            ))}
          </ul>
        )}
    </div>

    {data.lastQuery && (
      <div className={styles.diagBlock}>
        <div className={styles.diagSubheading}>Last query — {data.lastRowCount} row(s)</div>
        <code className={styles.diagQuery}>{data.lastQuery}</code>
      </div>
    )}
  </div>
  );
};

const ProcessDetails: React.FC<IProcessDetailsProps> = (props) => {
  const {
    selectedStep,
    isLoading,
    error,
    notFoundCode,
    openSubStepCode,
    onSubStepToggle,
    onRetry,
    title = DEFAULT_TITLE,
    headerIcon = DEFAULT_ICON,
    showBreadcrumb = true,
    singleOpenSubStep = true,
    emptyIcon = DEFAULT_EMPTY_ICON,
    emptyHeading = DEFAULT_EMPTY_HEADING,
    emptyBody = DEFAULT_EMPTY_BODY,
    diagnostics,
  } = props;

  // Controlled whenever the host supplies openSubStepCode at all — the prop is
  // then the source of truth and local state is ignored.
  const isControlled = openSubStepCode !== undefined;

  const [openCodes, setOpenCodes] = React.useState<string[]>([]);
  const instanceId = React.useMemo(() => `pd${Math.random().toString(36).slice(2, 9)}`, []);

  // Changing the selected step collapses every L4 row.
  const stepCode = selectedStep ? selectedStep.code : null;
  React.useEffect(() => { setOpenCodes([]); }, [stepCode]);

  const openNow: string[] = isControlled
    ? (openSubStepCode ? [openSubStepCode] : [])
    : openCodes;

  const handleToggle = (code: string): void => {
    const wasOpen = openNow.indexOf(code) !== -1;
    let next: string[];
    if (wasOpen) {
      next = singleOpenSubStep ? [] : openNow.filter((c) => c !== code);
    } else {
      next = singleOpenSubStep ? [code] : openNow.concat([code]);
    }
    if (!isControlled) { setOpenCodes(next); }
    if (onSubStepToggle) { onSubStepToggle(wasOpen ? null : code); }
  };

  const renderSubStep = (sub: ISubStep): JSX.Element => {
    const isOpen  = openNow.indexOf(sub.code) !== -1;
    const btnId   = `${instanceId}-btn-${slug(sub.code)}`;
    const panelId = `${instanceId}-pnl-${slug(sub.code)}`;

    return (
      <div className={styles.l4Row} key={sub.code}>
        <button
          type="button"
          id={btnId}
          className={`${styles.l4Header} ${isOpen ? styles.l4HeaderOpen : ''}`}
          aria-expanded={isOpen}
          aria-controls={panelId}
          onClick={() => handleToggle(sub.code)}
        >
          <span className={styles.l4Badge}>{sub.code}</span>
          <span className={styles.l4Name}>{sub.title}</span>
          <span className={`${styles.chevron} ${isOpen ? styles.chevronOpen : ''}`} aria-hidden="true">⌃</span>
        </button>

        {isOpen && (
          <div className={styles.l4Body} id={panelId} role="region" aria-labelledby={btnId}>
            <div className={styles.sectionLabel}>Outputs</div>
            <OutputList items={sub.outputs} />
          </div>
        )}
      </div>
    );
  };

  const renderBody = (): JSX.Element => {
    if (isLoading) {
      return (
        <div className={styles.loading} aria-busy="true" aria-label="Loading process details">
          <div className={`${styles.shimmer} ${styles.shimmerA}`} />
          <div className={`${styles.shimmer} ${styles.shimmerB}`} />
          <div className={`${styles.shimmer} ${styles.shimmerC}`} />
        </div>
      );
    }

    if (error) {
      return (
        <div className={styles.empty}>
          <div className={styles.emptyHeading}>Couldn&rsquo;t load process details</div>
          <div className={styles.emptyBody}>{error}</div>
          {onRetry && (
            <button type="button" className={styles.retryBtn} onClick={onRetry}>Retry</button>
          )}
        </div>
      );
    }

    if (!selectedStep) {
      // A selected code with no rows is a different message from no selection —
      // the first means "nobody has written this up yet", the second "click one".
      return notFoundCode ? (
        <div className={styles.empty}>
          <div className={styles.emptyIcon} aria-hidden="true">📄</div>
          <div className={styles.emptyHeading}>No details recorded for {notFoundCode}</div>
          <div className={styles.emptyBody}>
            This process has no row in the process details list yet.
          </div>
        </div>
      ) : (
        <div className={styles.empty}>
          <div className={styles.emptyIcon} aria-hidden="true">{emptyIcon}</div>
          <div className={styles.emptyHeading}>{emptyHeading}</div>
          <div className={styles.emptyBody}>{emptyBody}</div>
        </div>
      );
    }

    const hasSubSteps = selectedStep.subSteps && selectedStep.subSteps.length > 0;

    return (
      <>
        <div className={styles.titleBlock}>
          {showBreadcrumb && (
            <div className={styles.breadcrumb}>
              {selectedStep.code}
              {selectedStep.breadcrumb ? ` · ${selectedStep.breadcrumb}` : ''}
            </div>
          )}
          <h3 className={styles.stepTitle}>{selectedStep.title}</h3>
        </div>

        {selectedStep.purpose && (
          <div className={styles.section}>
            <div className={styles.sectionLabel}>Purpose</div>
            <p className={styles.purpose}>{selectedStep.purpose}</p>
          </div>
        )}

        {selectedStep.inputs.length > 0 && (
          <div className={styles.section}>
            <div className={styles.sectionLabel}>Inputs</div>
            <ul className={styles.list}>
              {selectedStep.inputs.map((input, i) => <li key={`${input}_${i}`}>{input}</li>)}
            </ul>
          </div>
        )}

        {/* With L4 sub-steps the flat OUTPUTS section is dropped — outputs live
            inside each expanded row instead. */}
        {hasSubSteps ? (
          <div className={styles.l4Block}>
            {selectedStep.subSteps.map(renderSubStep)}
          </div>
        ) : (
          <div className={styles.section}>
            <div className={styles.sectionLabel}>Outputs</div>
            <OutputList items={selectedStep.outputs} />
          </div>
        )}
      </>
    );
  };

  return (
    <div className={styles.panel} role="region" aria-label={title || DEFAULT_TITLE}>
      <div className={styles.header}>
        {headerIcon && <span className={styles.headerIcon} aria-hidden="true">{headerIcon}</span>}
        <span className={styles.headerTitle}>{title || DEFAULT_TITLE}</span>
      </div>
      <div className={styles.body}>
        {renderBody()}
        {diagnostics && <Diagnostics data={diagnostics} />}
      </div>
    </div>
  );
};

export default ProcessDetails;
