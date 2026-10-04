import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { MessageBarButton } from '@fluentui/react/lib/Button';
import styles from './MasterRasci.module.scss';
import {
  IMasterRasciProps, IMasterRasciData, IFilterState, RasciView,
} from './IMasterRasciProps';
import { EMPTY_FILTERS, filterRows } from './masterRasciModel';
import FilterPanel from './FilterPanel';
import ActivityListView from './ActivityListView';
import MatrixView from './MatrixView';

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

// ── Shareable filters ───────────────────────────────────────────────────────
// ?phase=&group=&role=&q=&view= — the role travels by name, not by index,
// because indexes shift whenever the list gains a role.

const PARAMS = ['phase', 'group', 'role', 'q', 'view'];

const readUrl = (): { filters: Partial<IFilterState> & { roleName?: string }; view?: RasciView } => {
  try {
    const p = new URLSearchParams(window.location.search);
    const view = p.get('view');
    return {
      filters: {
        phase: p.get('phase') || undefined,
        group: p.get('group') || undefined,
        roleName: p.get('role') || undefined,
        q: p.get('q') || undefined,
      },
      view: view === 'matrix' || view === 'list' ? view : undefined,
    };
  } catch {
    return { filters: {} };
  }
};

const writeUrl = (data: IMasterRasciData, f: IFilterState, view: RasciView, defaultView: RasciView): void => {
  try {
    const p = new URLSearchParams(window.location.search);
    PARAMS.forEach((k) => p.delete(k));
    if (f.phase !== 'All') { p.set('phase', f.phase); }
    if (f.group !== 'All') { p.set('group', f.group); }
    if (f.role >= 0 && data.roles[f.role]) { p.set('role', data.roles[f.role].name); }
    if (f.q) { p.set('q', f.q); }
    if (view !== defaultView) { p.set('view', view); }
    const qs = p.toString();
    const next = window.location.pathname + (qs ? `?${qs}` : '') + window.location.hash;
    if (next !== window.location.pathname + window.location.search + window.location.hash) {
      window.history.replaceState(window.history.state, '', next);
    }
  } catch { /* history is unavailable in some hosts; sharing is a nicety, not a need */ }
};

/** Turns URL values into a state that is valid for this data, dropping anything stale. */
const resolveInitial = (
  data: IMasterRasciData,
  fromUrl: ReturnType<typeof readUrl>['filters']
): IFilterState => {
  const phases: { [p: string]: true } = {};
  const groups: { [g: string]: string } = {};
  data.rows.forEach((r) => { phases[r.phase] = true; groups[r.group] = r.phase; });

  let role = -1;
  if (fromUrl.roleName) {
    const want = fromUrl.roleName.toLowerCase();
    data.roles.forEach((r, i) => { if (r.name.toLowerCase() === want) { role = i; } });
  }
  const group = fromUrl.group && groups[fromUrl.group] !== undefined ? fromUrl.group : 'All';
  const phaseWanted = fromUrl.phase && phases[fromUrl.phase] ? fromUrl.phase : undefined;
  return {
    // A group implies its phase; an explicit phase that contradicts it loses.
    phase: group !== 'All' ? groups[group] : (phaseWanted || 'All'),
    group,
    role,
    q: fromUrl.q || '',
  };
};

// ── Component ───────────────────────────────────────────────────────────────

const MasterRasci: React.FC<IMasterRasciProps> = ({
  load, listTitle, documentRef, defaultView, collapseAfter, labels,
}) => {
  const [data, setData] = React.useState<IMasterRasciData | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [attempt, setAttempt] = React.useState(0);
  const [filters, setFilters] = React.useState<IFilterState>(EMPTY_FILTERS);
  const [activityWidth, setActivityWidth] = React.useState<number | undefined>(undefined);
  const [view, setView] = React.useState<RasciView>(readUrl().view || defaultView);

  React.useEffect(() => {
    let cancelled = false;
    setData(undefined);
    setError(undefined);
    load().then(
      (d) => {
        if (cancelled) { return; }
        setFilters(resolveInitial(d, readUrl().filters));
        setData(d);
      },
      (err: unknown) => {
        if (cancelled) { return; }
        setError(err instanceof Error ? err.message : String(err));
      }
    );
    return () => { cancelled = true; };
  }, [load, attempt]);

  // The page's own default view is the fallback until a reader picks one.
  const userPickedView = React.useRef(false);
  React.useEffect(() => {
    if (!userPickedView.current && !readUrl().view) { setView(defaultView); }
  }, [defaultView]);

  React.useEffect(() => {
    if (data) { writeUrl(data, filters, view, defaultView); }
  }, [data, filters, view, defaultView]);

  const rows = React.useMemo(() => (data ? filterRows(data, filters) : []), [data, filters]);

  const pickView = (v: RasciView): void => {
    userPickedView.current = true;
    setView(v);
  };

  if (error) {
    return (
      <div className={styles.root}>
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline
          actions={<MessageBarButton onClick={() => setAttempt((n) => n + 1)}>Retry</MessageBarButton>}
        >
          Couldn&rsquo;t load the Master RASCI. {error}
        </MessageBar>
      </div>
    );
  }

  if (!data) {
    return (
      <div className={styles.root}>
        <div className={cx(styles.panel, styles.results)} aria-busy="true" aria-label="Loading the Master RASCI">
          <div className={styles.resultsHead}>
            <div className={cx(styles.resultsCount, styles.resultsMuted)}>Loading processes&hellip;</div>
          </div>
          <div className={styles.skeleton}>
            {[0, 1, 2, 3, 4, 5].map((i) => <div key={i} className={styles.skeletonRow} />)}
          </div>
        </div>
      </div>
    );
  }

  const role = filters.role;
  const roleNote = role >= 0 && data.roles[role] ? `  ·  viewing as ${data.roles[role].name}` : '';

  return (
    <div className={styles.root}>
      <FilterPanel data={data} filters={filters} onChange={setFilters} labels={labels} />

      <div className={cx(styles.panel, styles.results)}>
        <div className={styles.resultsHead}>
          <div className={styles.resultsCount} aria-live="polite">
            <span style={{ fontWeight: 700 }}>{rows.length}</span>{' '}
            <span className={styles.resultsMuted}>of {data.rows.length} processes</span>
            <span className={styles.resultsMuted}>{roleNote}</span>
          </div>
          <div className={styles.viewToggle} role="group" aria-label="View">
            <button
              type="button"
              aria-pressed={view === 'list'}
              className={cx(styles.viewBtn, view === 'list' && styles.viewBtnOn)}
              onClick={() => pickView('list')}
            >{labels.viewList}</button>
            <button
              type="button"
              aria-pressed={view === 'matrix'}
              className={cx(styles.viewBtn, view === 'matrix' && styles.viewBtnOn)}
              onClick={() => pickView('matrix')}
            >{labels.viewMatrix}</button>
          </div>
        </div>

        {rows.length === 0 && (
          <div className={styles.empty}>
            {data.rows.length === 0
              ? `The “${listTitle}” list has no processes yet.`
              : 'No processes match these filters.'}
          </div>
        )}
        {rows.length > 0 && view === 'list' && (
          <ActivityListView data={data} rows={rows} role={role} collapseAfter={collapseAfter}
            activityWidth={activityWidth}
            onActivityWidthChange={setActivityWidth}
          />
        )}
        {rows.length > 0 && view === 'matrix' && (
          <MatrixView
            data={data}
            rows={rows}
            role={role}
            activityWidth={activityWidth}
            onActivityWidthChange={setActivityWidth}
          />
        )}
      </div>

      <div className={styles.footer}>
        {documentRef ? `${documentRef} · ` : ''}Data source: SharePoint list &ldquo;{listTitle}&rdquo;. A/R = Accountable and Responsible.
      </div>
    </div>
  );
};

export default MasterRasci;
