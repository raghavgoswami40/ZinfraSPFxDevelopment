import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { MessageBarButton } from '@fluentui/react/lib/Button';
import styles from './MasterRasci.module.scss';
import {
  IMasterRasciProps, IMasterRasciData, IFilterState, RasciView, ChipCode,
} from './IMasterRasciProps';
import { EMPTY_FILTERS, filterRows, hasActiveFilters, phaseKey } from './masterRasciModel';
import { SearchIcon, ListIcon, MatrixIcon } from './icons';
import FilterPanel from './FilterPanel';
import ActivityListView from './ActivityListView';
import MatrixView from './MatrixView';
import { ThemeContext, ThemeName, readStoredTheme, storeTheme } from './theme';
import { ensureOutfitFont } from '../../../shared/outfitFont';
import PageBanner from '../../pageBanner/components/PageBanner';

// eslint-disable-next-line @typescript-eslint/no-var-requires
const logoOnDark: string = require('../../pageBanner/assets/zinfra-logo.png');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const logoOnLight: string = require('../../pageBanner/assets/zinfra-logo-dark.png');

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

// ── Shareable filters ───────────────────────────────────────────────────────
// ?phase=&group=&role=&code=&q=&view= — the role travels by name, not by index,
// because indexes shift whenever the list gains a role.

const PARAMS = ['phase', 'group', 'role', 'code', 'q', 'view'];
const VALID_CODES = ['A', 'R', 'S', 'C', 'I'];

const readUrl = (): { filters: Partial<IFilterState> & { roleName?: string }; view?: RasciView } => {
  try {
    const p = new URLSearchParams(window.location.search);
    const view = p.get('view');
    const code = (p.get('code') || '').toUpperCase();
    return {
      filters: {
        phase: p.get('phase') || undefined,
        group: p.get('group') || undefined,
        roleName: p.get('role') || undefined,
        code: VALID_CODES.indexOf(code) !== -1 ? (code as ChipCode) : undefined,
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
    if (f.code) { p.set('code', f.code); }
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
    code: fromUrl.code || '',
    q: fromUrl.q || '',
  };
};

// ── Component ───────────────────────────────────────────────────────────────

const MasterRasciBody: React.FC<IMasterRasciProps> = ({
  load, listTitle, defaultView, collapseAfter, labels,
}) => {
  const [data, setData] = React.useState<IMasterRasciData | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [attempt, setAttempt] = React.useState(0);
  const [filters, setFilters] = React.useState<IFilterState>(EMPTY_FILTERS);
  const [activityWidth, setActivityWidth] = React.useState<number | undefined>(undefined);
  const [collapsed, setCollapsed] = React.useState<{ [group: string]: boolean }>({});
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

  // Groups present in the current result, for the collapse-all toggle.
  const groupNames = React.useMemo(() => {
    const seen: string[] = [];
    rows.forEach((r) => { if (seen.indexOf(r.group) === -1) { seen.push(r.group); } });
    return seen;
  }, [rows]);
  const phaseNames = React.useMemo(() => {
    const seen: string[] = [];
    rows.forEach((r) => { if (seen.indexOf(r.phase) === -1) { seen.push(r.phase); } });
    return seen;
  }, [rows]);
  // "Collapse all" folds every phase and every L2 group, so opening one phase afterwards shows
  // its groups still folded. "Expand all" opens all of them again.
  const allCollapsed = phaseNames.length > 0 && phaseNames.every((p) => collapsed[phaseKey(p)]);

  const toggleGroup = (group: string): void =>
    setCollapsed((prev) => ({ ...prev, [group]: !prev[group] }));
  const toggleAll = (): void =>
    setCollapsed((prev) => {
      const next = { ...prev };
      phaseNames.forEach((p) => { next[phaseKey(p)] = !allCollapsed; });
      groupNames.forEach((g) => { next[g] = !allCollapsed; });
      return next;
    });

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
        <div className={styles.card} aria-busy="true" aria-label="Loading the Master RASCI">
          <div className={styles.bar}>
            <div className={cx(styles.count, styles.muted)}>Loading processes&hellip;</div>
          </div>
          <div className={styles.skeleton}>
            {[0, 1, 2, 3, 4, 5].map((i) => <div key={i} className={styles.skeletonRow} />)}
          </div>
        </div>
      </div>
    );
  }

  const role = filters.role;
  const roleName = role >= 0 && data.roles[role] ? data.roles[role].name : '';
  const filtered = hasActiveFilters(filters);

  return (
    <div className={styles.root}>
      <section className={styles.card}>
        <FilterPanel data={data} filters={filters} onChange={setFilters} labels={labels} />

        <div className={styles.bar}>
          <div className={styles.count} aria-live="polite">
            <strong>{rows.length}</strong>
            <span className={styles.muted}> of {data.rows.length} processes</span>
            {roleName && <span className={styles.rolePill}>Viewing as {roleName}</span>}
          </div>
          <div className={styles.seg} role="group" aria-label="View">
            <button
              type="button"
              aria-pressed={view === 'list'}
              className={cx(styles.segBtn, view === 'list' && styles.segBtnOn)}
              onClick={() => pickView('list')}
            ><ListIcon />{labels.viewList}</button>
            <button
              type="button"
              aria-pressed={view === 'matrix'}
              className={cx(styles.segBtn, view === 'matrix' && styles.segBtnOn)}
              onClick={() => pickView('matrix')}
            ><MatrixIcon />{labels.viewMatrix}</button>
          </div>
        </div>

        {rows.length === 0 && (
          <div className={styles.empty}>
            <div className={styles.emptyIcon}><SearchIcon /></div>
            {data.rows.length === 0 ? (
              <div className={styles.emptyTitle}>The &ldquo;{listTitle}&rdquo; list has no processes yet</div>
            ) : (
              <>
                <div className={styles.emptyTitle}>No processes match these filters</div>
                <div className={styles.emptyBody}>Remove a filter or try a broader search.</div>
                {filtered && (
                  <button type="button" className={styles.emptyBtn} onClick={() => setFilters(EMPTY_FILTERS)}>
                    Clear all filters
                  </button>
                )}
              </>
            )}
          </div>
        )}
        {rows.length > 0 && view === 'list' && (
          <ActivityListView
            data={data}
            rows={rows}
            role={role}
            collapseAfter={collapseAfter}
            activityWidth={activityWidth}
            onActivityWidthChange={setActivityWidth}
            collapsed={collapsed}
            onToggleGroup={toggleGroup}
            onToggleAll={toggleAll}
            allCollapsed={allCollapsed}
          />
        )}
        {rows.length > 0 && view === 'matrix' && (
          <MatrixView
            data={data}
            rows={rows}
            role={role}
            activityWidth={activityWidth}
            onActivityWidthChange={setActivityWidth}
            collapsed={collapsed}
            onToggleGroup={toggleGroup}
            onToggleAll={toggleAll}
            allCollapsed={allCollapsed}
          />
        )}
      </section>

    </div>
  );
};

/** The wrapper is the size container the scale unit is measured against. */
const MasterRasci: React.FC<IMasterRasciProps> = (props) => {
  React.useEffect(ensureOutfitFont, []);
  const [theme, setThemeState] = React.useState<ThemeName>(readStoredTheme);
  const value = React.useMemo(() => ({
    theme,
    setTheme: (t: ThemeName): void => { storeTheme(t); setThemeState(t); },
  }), [theme]);

  return (
    <div className={styles.scope} data-mr-scope="" data-theme={theme}>
      <PageBanner
        title={props.bannerTitle === undefined ? 'Master RASCI' : props.bannerTitle}
        height={props.showTitle !== false ? 130 : 56}
        showTitle={props.showTitle !== false}
        backgroundColor="#000000"
        showSparkles={props.showSparkles !== false}
        followPageTheme={false}
        theme={theme}
        embedded
        showLogo={false}
      />
      <div className={styles.content}>
        <ThemeContext.Provider value={value}>
          <MasterRasciBody {...props} />
        </ThemeContext.Provider>
        <div className={styles.logoRow}>
          <img
            className={styles.logo}
            src={theme === 'light' ? logoOnLight : logoOnDark}
            alt="Zinfra - Your trusted partner in energy"
          />
        </div>
      </div>
    </div>
  );
};

export default MasterRasci;
