import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './ProcessHierarchy.module.scss';
import { StoredPhase, StoredSection, StoredStep } from '../../verticalFlow/components/IVerticalFlowProps';
import { buildCodeMap, getShape, StepShape } from '../../verticalFlow/stepSelection';
import { ensureOutfitFont } from '../../../shared/outfitFont';
import PageBanner from '../../pageBanner/components/PageBanner';
import ProcessPanelContainer, { IPanelConfig } from './ProcessPanelContainer';

export interface IProcessHierarchyProps {
  phases: StoredPhase[];
  spHttpClient: SPHttpClient;
  webUrl: string;
  panelConfig: IPanelConfig;
  /** Page is in edit mode: titles, steps and sections can be changed. */
  bannerTitle: string;
  showSparkles: boolean;
  isEditMode: boolean;
  onPhasesChange: (phases: StoredPhase[]) => void;
}

// eslint-disable-next-line @typescript-eslint/no-var-requires
const logoOnDark: string = require('../../pageBanner/assets/zinfra-logo.png');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const logoOnLight: string = require('../../pageBanner/assets/zinfra-logo-dark.png');

const SHAPE_ANNOUNCE: Record<string, string> = {
  box: '', approval: ' (client approval)', sapStatus: ' (SAP status)', gate: ' (gate)', group: '',
};

// "5 CONTROL" is drawn as a horizontal band; every other phase is vertical columns.
const isControl = (p: StoredPhase): boolean => /^\s*5\b/.test(p.title);
const uid = (prefix: string): string => `${prefix}_${Date.now()}_${Math.floor(Math.random() * 9999)}`;

type EditTarget = { id: string; type: 'step' | 'section' | 'phase' };
interface StepCtx { phaseId: string; sectionId: string }

// What the add-step picker offers, and the starting text of each new step.
const SHAPE_OPTIONS: Array<{ shape: StepShape; menuLabel: string; defaultLabel: string; glyph: string }> = [
  { shape: 'box',       menuLabel: 'Box',             defaultLabel: 'New step',        glyph: 'pvBox' },
  { shape: 'approval',  menuLabel: 'Client approval', defaultLabel: 'Client Approval', glyph: 'pvApproval' },
  { shape: 'sapStatus', menuLabel: 'SAP status',      defaultLabel: 'STAT',            glyph: 'pvSap' },
  { shape: 'gate',      menuLabel: 'Gate',            defaultLabel: 'Gate',            glyph: 'pvGate' },
  { shape: 'group',     menuLabel: 'Dashed group',    defaultLabel: 'Group title',     glyph: 'pvGroup' },
];

type Theme = 'dark' | 'light';
const THEME_KEY = 'ph-theme';

// localStorage can throw (private windows, blocked site data), so every access
// is guarded and the page works without it.
const readTheme = (): Theme => {
  try { return window.localStorage.getItem(THEME_KEY) === 'light' ? 'light' : 'dark'; } catch { return 'dark'; }
};

const SearchIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true" focusable="false">
    <circle cx="11" cy="11" r="7" /><path d="M20 20l-3.5-3.5" />
  </svg>
);

const SunIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true" focusable="false">
    <circle cx="12" cy="12" r="4" />
    <path d="M12 2v2M12 20v2M2 12h2M20 12h2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4" />
  </svg>
);

const MoonIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
    <path d="M21 12.8A9 9 0 1111.2 3a7 7 0 009.8 9.8z" />
  </svg>
);

const tidy = (s: string): string => (s || '').replace(/\s*\n\s*/g, ' ').replace(/\s{2,}/g, ' ').trim();

const ProcessHierarchy: React.FC<IProcessHierarchyProps> = ({ phases, spHttpClient, webUrl, panelConfig, bannerTitle, showSparkles, isEditMode, onPhasesChange }) => {
  const [selected, setSelected] = React.useState<{ id: string; code: string } | undefined>();
  const [query, setQuery] = React.useState('');
  React.useEffect(ensureOutfitFont, []);
  const [theme, setTheme] = React.useState<Theme>(readTheme);
  const toggleTheme = (): void => {
    const next: Theme = theme === 'dark' ? 'light' : 'dark';
    setTheme(next);
    try { window.localStorage.setItem(THEME_KEY, next); } catch { /* theme still applies for this visit */ }
  };
  const lastFocus = React.useRef<HTMLElement | null>(null);
  const closeRef = React.useRef<HTMLButtonElement>(null);
  const phasesRef = React.useRef<HTMLDivElement>(null);
  const [l3Width, setL3Width] = React.useState<number | undefined>();
  const scopeRef = React.useRef<HTMLDivElement>(null);
  // Where the popup pins while the page scrolls: just below SharePoint's top bar,
  // measured inside whichever element the page scrolls.
  const [drawerBox, setDrawerBox] = React.useState({ stickyTop: 0, visibleTop: 48 });

  // Control's L3 boxes copy the width of the L3 boxes in the phases above, so
  // they stay the same size whatever the screen width does to the columns.
  React.useLayoutEffect(() => {
    const el = phasesRef.current;
    if (!el) return undefined;
    const measure = (): void => {
      const box = el.querySelector('[data-l3]');
      if (box) {
        const w = Math.round(box.getBoundingClientRect().width);
        setL3Width(prev => (prev === w ? prev : w));
      }
    };
    measure();
    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', measure);
      return () => window.removeEventListener('resize', measure);
    }
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [phases]);

  // The popup opens level with the top of this web part and, as the page scrolls,
  // stays just below SharePoint's own top bar instead of sliding behind it.
  React.useLayoutEffect(() => {
    if (!selected) return undefined;
    const findScroller = (start: HTMLElement | null): HTMLElement | undefined => {
      let el = start ? start.parentElement : null;
      while (el && el !== document.body) {
        const oy = window.getComputedStyle(el).overflowY;
        if ((oy === 'auto' || oy === 'scroll') && el.scrollHeight > el.clientHeight) return el;
        el = el.parentElement;
      }
      return undefined;
    };
    const measure = (): void => {
      const bar = document.getElementById('SuiteNavWrapper');
      const barBottom = bar ? Math.round(bar.getBoundingClientRect().bottom) : 48;
      const scroller = findScroller(scopeRef.current);
      const scrollerTop = scroller ? Math.round(scroller.getBoundingClientRect().top) : 0;
      const visibleTop = Math.max(barBottom > 0 ? barBottom : 48, scrollerTop);
      setDrawerBox({ stickyTop: Math.max(0, visibleTop - scrollerTop), visibleTop });
    };
    measure();
    window.addEventListener('resize', measure);
    return () => window.removeEventListener('resize', measure);
  }, [selected]);

  const close = React.useCallback((): void => {
    setSelected(undefined);
    if (lastFocus.current) { lastFocus.current.focus(); }
  }, []);

  React.useEffect(() => {
    if (!selected) return undefined;
    if (closeRef.current) closeRef.current.focus();
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') close(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [selected, close]);

  const q = query.trim().toLowerCase();

  // ── Edit mode ─────────────────────────────────────────────────────────
  // The same authoring actions as the Vertical Flow web part: rename any title
  // or label by clicking it, add steps (box, client approval, SAP status, gate
  // or dashed group), add steps inside a group, add sections, and remove any of
  // them. Step codes are generated, so they renumber on their own.
  const [editing, setEditing] = React.useState<EditTarget | undefined>();
  const [editValue, setEditValue] = React.useState('');
  const [addMenuId, setAddMenuId] = React.useState<string | undefined>();

  React.useEffect(() => {
    if (!addMenuId) return undefined;
    const onDown = (e: MouseEvent): void => {
      const t = e.target as HTMLElement | null;
      if (!t || !t.closest || !t.closest('[data-addmenu]')) setAddMenuId(undefined);
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [addMenuId]);

  const startEdit = (id: string, type: EditTarget['type'], current: string): void => {
    setEditing({ id, type });
    setEditValue(current);
  };
  const cancelEdit = (): void => setEditing(undefined);

  const commitEdit = (): void => {
    if (!editing) return;
    const { id, type } = editing;
    const val = editValue.trim() || '(untitled)';
    // Recurses so steps nested inside a dashed group are renamed too.
    const renameStep = (st: StoredStep): StoredStep => {
      let next = st.id === id && type === 'step' ? { ...st, label: val } : st;
      if (next.children) next = { ...next, children: next.children.map(renameStep) };
      return next;
    };
    onPhasesChange(phases.map(ph => {
      if (type === 'phase' && ph.id === id) return { ...ph, title: val };
      return {
        ...ph,
        sections: ph.sections.map(sec => {
          if (type === 'section' && sec.id === id) return { ...sec, title: val };
          return { ...sec, steps: sec.steps.map(renameStep) };
        }),
      };
    }));
    setEditing(undefined);
  };

  const mapSection = (phaseId: string, sectionId: string, fn: (s: StoredSection) => StoredSection): void =>
    onPhasesChange(phases.map(ph => ph.id !== phaseId ? ph : {
      ...ph,
      sections: ph.sections.map(sec => sec.id !== sectionId ? sec : fn(sec)),
    }));

  const addSection = (phaseId: string): void => {
    const id = uid(`${phaseId}_sec`);
    onPhasesChange(phases.map(ph => ph.id !== phaseId ? ph : {
      ...ph,
      sections: [...ph.sections, { id, title: 'New section', steps: [{ id: uid(`${id}_step`), label: 'New step' }] }],
    }));
  };

  const removeSection = (phaseId: string, sectionId: string): void =>
    onPhasesChange(phases.map(ph => ph.id !== phaseId ? ph : {
      ...ph, sections: ph.sections.filter(s => s.id !== sectionId),
    }));

  const addStep = (phaseId: string, sectionId: string, shape: StepShape): void => {
    const id = uid(`${sectionId}_step`);
    const label = SHAPE_OPTIONS.filter(o => o.shape === shape)[0].defaultLabel;
    // A new dashed group starts empty; its steps are added from the group itself.
    const step: StoredStep = shape === 'group' ? { id, label, shape, children: [] } : { id, label, shape };
    mapSection(phaseId, sectionId, sec => ({ ...sec, steps: [...sec.steps, step] }));
    setAddMenuId(undefined);
  };

  // Groups cannot nest, so 'group' is not offered inside one.
  const addChildStep = (phaseId: string, sectionId: string, groupId: string, shape: StepShape): void => {
    const label = SHAPE_OPTIONS.filter(o => o.shape === shape)[0].defaultLabel;
    const child: StoredStep = { id: uid(`${groupId}_child`), label, shape };
    mapSection(phaseId, sectionId, sec => ({
      ...sec,
      steps: sec.steps.map(st => st.id !== groupId ? st : { ...st, children: [...(st.children || []), child] }),
    }));
  };

  const dropStep = (steps: StoredStep[], stepId: string): StoredStep[] =>
    steps.filter(st => st.id !== stepId).map(st => st.children ? { ...st, children: dropStep(st.children, stepId) } : st);

  const removeStep = (phaseId: string, sectionId: string, stepId: string): void =>
    mapSection(phaseId, sectionId, sec => ({ ...sec, steps: dropStep(sec.steps, stepId) }));

  // A label that turns into a text box when clicked. Enter saves, Esc cancels.
  const renderEditableLabel = (step: StoredStep, className: string, editorClassName: string): JSX.Element => {
    if (editing && editing.id === step.id && editing.type === 'step') {
      return (
        <textarea
          className={editorClassName}
          value={editValue}
          autoFocus
          aria-label="Step name"
          onChange={e => setEditValue(e.target.value)}
          onBlur={commitEdit}
          onKeyDown={e => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitEdit(); }
            if (e.key === 'Escape') cancelEdit();
          }}
        />
      );
    }
    return (
      <span className={`${className} ${styles.editable}`} title="Click to rename" onClick={() => startEdit(step.id, 'step', step.label)}>
        {step.label}
      </span>
    );
  };

  // A phase or section title that turns into a one-line input when clicked.
  const renderEditableTitle = (id: string, type: 'phase' | 'section', text: string, emptyHint: string): JSX.Element => {
    if (editing && editing.id === id && editing.type === type) {
      return (
        <input
          className={styles.titleEditor}
          value={editValue}
          autoFocus
          aria-label="Title"
          onChange={e => setEditValue(e.target.value)}
          onBlur={commitEdit}
          onKeyDown={e => {
            if (e.key === 'Enter') { e.preventDefault(); commitEdit(); }
            if (e.key === 'Escape') cancelEdit();
          }}
        />
      );
    }
    return (
      <span className={styles.editable} title="Click to rename" onClick={() => startEdit(id, type, text)}>
        {text || emptyHint}
      </span>
    );
  };

  // The "+ Add step" button and its shape picker, shown at the foot of a section.
  const renderAddStep = (phaseId: string, sectionId: string): JSX.Element => (
    <div className={styles.addWrap} data-addmenu="">
      <button type="button" className={styles.addBtn} aria-expanded={addMenuId === sectionId}
        onClick={() => setAddMenuId(addMenuId === sectionId ? undefined : sectionId)}>
        + Add step
      </button>
      {addMenuId === sectionId && (
        <div className={styles.addMenu} role="menu">
          {SHAPE_OPTIONS.map(o => (
            <button key={o.shape} type="button" role="menuitem" className={styles.addMenuItem} onClick={() => addStep(phaseId, sectionId, o.shape)}>
              <span className={`${styles.glyph} ${(styles as Record<string, string>)[o.glyph]}`} aria-hidden="true" />
              {o.menuLabel}
            </button>
          ))}
        </div>
      )}
    </div>
  );

  const renderStep = (step: StoredStep, codes: Map<string, string>, ctx: StepCtx): JSX.Element => {
    const shape = getShape(step);
    const code = codes.get(step.id);
    const dim = !!q && `${code || ''} ${step.label}`.toLowerCase().indexOf(q) === -1;
    const removeBtn = (
      <button type="button" className={styles.removeBtn} title="Remove" aria-label={`Remove ${step.label}`}
        onClick={() => removeStep(ctx.phaseId, ctx.sectionId, step.id)}>×</button>
    );

    if (shape === 'group') {
      return (
        <div key={step.id} className={`${styles.shapeGroup}${dim ? ' ' + styles.stepDim : ''}`}>
          {isEditMode
            ? renderEditableLabel(step, styles.groupTitle, styles.groupTitleEditor)
            : <span className={styles.groupTitle}>{step.label}</span>}
          {(step.children || []).map(c => renderStep(c, codes, ctx))}
          {isEditMode && (
            <div className={styles.groupAdd}>
              {SHAPE_OPTIONS.filter(o => o.shape !== 'group').map(o => (
                <button key={o.shape} type="button" className={`${styles.groupAddBtn} ${(styles as Record<string, string>)[o.glyph]}`}
                  title={`Add ${o.menuLabel} to this group`} aria-label={`Add ${o.menuLabel} to this group`}
                  onClick={() => addChildStep(ctx.phaseId, ctx.sectionId, step.id, o.shape)} />
              ))}
            </div>
          )}
          {isEditMode && removeBtn}
        </div>
      );
    }

    const cls = [
      styles.step,
      shape === 'approval' ? styles.shapeApproval : '',
      shape === 'sapStatus' ? styles.shapeSapStatus : '',
      shape === 'gate' ? styles.shapeGate : '',
      !isEditMode && shape === 'box' && !!code ? styles.stepClickable : '',
      !isEditMode && selected && selected.id === step.id ? styles.stepSelected : '',
      dim ? styles.stepDim : '',
    ].filter(Boolean).join(' ');

    if (isEditMode) {
      return (
        <div key={step.id} className={cls} data-l3={shape === 'box' ? '' : undefined}>
          {code && <span className={styles.stepCodeBadge}>{code}</span>}
          {shape === 'sapStatus' && !step.noPrefix && <span className={styles.sapPrefix}>SAP Status</span>}
          {renderEditableLabel(step, styles.stepLabel, styles.stepEditor)}
          {removeBtn}
        </div>
      );
    }

    if (shape === 'gate') {
      return (
        <div key={step.id} className={cls} role="img" aria-label={`${step.label}${SHAPE_ANNOUNCE.gate}`}>
          <span className={styles.stepLabel}>{step.label}</span>
        </div>
      );
    }

    const clickable = shape === 'box' && !!code;
    const inner = (
      <>
        {code && <span className={styles.stepCodeBadge}>{code}</span>}
        {shape === 'sapStatus' && !step.noPrefix && <span className={styles.sapPrefix}>SAP Status</span>}
        <span className={styles.stepLabel}>{step.label}</span>
      </>
    );

    if (clickable) {
      return (
        <button
          key={step.id}
          type="button"
          data-l3=""
          className={cls}
          aria-label={`${code} — ${step.label}`}
          aria-pressed={!!selected && selected.id === step.id}
          onClick={(e) => {
            lastFocus.current = e.currentTarget;
            setSelected(prev => (prev && prev.id === step.id ? undefined : { id: step.id, code: code as string }));
          }}
        >
          {inner}
        </button>
      );
    }
    return (
      <div key={step.id} className={cls} role="img" aria-label={`${code ? code + ' — ' : ''}${step.label}${SHAPE_ANNOUNCE[shape]}`}>
        {inner}
      </div>
    );
  };

  const renderColumn = (section: StoredSection, phaseId: string): JSX.Element => {
    const codes = buildCodeMap(section);
    const title = tidy(section.title);
    const ctx: StepCtx = { phaseId, sectionId: section.id };
    return (
      <section key={section.id} className={`${styles.column}${section.variant === 'grey' ? ' ' + styles.columnGrey : ''}`} aria-label={title || undefined}>
        {(title || isEditMode) && (
          <h3 className={styles.columnHeader} style={{ margin: 0 }}>
            {isEditMode ? renderEditableTitle(section.id, 'section', title, '(no title)') : title}
          </h3>
        )}
        <div className={styles.columnBody} style={title || isEditMode ? undefined : { paddingTop: 52 }}>
          {section.steps.map(s => renderStep(s, codes, ctx))}
          {isEditMode && renderAddStep(phaseId, section.id)}
        </div>
        {isEditMode && (
          <button type="button" className={styles.removeSectionBtn} title="Remove section" aria-label={`Remove section ${title}`}
            onClick={() => removeSection(phaseId, section.id)}>×</button>
        )}
      </section>
    );
  };

  const vertical = phases.filter(p => !isControl(p));
  const control = phases.filter(isControl);

  return (
    <div className={styles.scope} data-theme={theme} ref={scopeRef}>
      <PageBanner
        title={bannerTitle}
        backgroundColor="#000000"
        showSparkles={showSparkles}
        followPageTheme={false}
        theme={theme}
        embedded
        showLogo={false}
      >
    <div className={styles.toolbar}>
    <ul className={styles.legend} aria-label="Shape key">
      <li><span className={`${styles.swatch} ${styles.swatchBox}`} aria-hidden="true" />Process step</li>
      <li><span className={`${styles.swatch} ${styles.swatchSap}`} aria-hidden="true" />SAP status</li>
      <li><span className={`${styles.swatch} ${styles.swatchPill}`} aria-hidden="true" />Client approval</li>
      <li><span className={`${styles.swatch} ${styles.swatchGate}`} aria-hidden="true" />Gate</li>
    </ul>
      <div className={styles.tools}>
        <label className={styles.searchWrap}>
          <span className={styles.searchIcon}><SearchIcon /></span>
          <input
            className={styles.search}
            type="search"
            placeholder="Find a process, e.g. design"
            aria-label="Find a process"
            value={query}
            onChange={e => setQuery(e.target.value)}
          />
        </label>
        <button
          type="button"
          role="switch"
          aria-checked={theme === 'dark'}
          aria-label="Dark theme"
          className={styles.themeBtn}
          onClick={toggleTheme}
        >
          <span className={styles.themeIcon}>{theme === 'dark' ? <MoonIcon /> : <SunIcon />}</span>
          <span className={styles.themeLabel}>{theme === 'dark' ? 'Dark' : 'Light'}</span>
          <span className={styles.themeTrack} aria-hidden="true"><span className={styles.themeKnob} /></span>
        </button>
      </div>
    </div>
      </PageBanner>

      <div className={styles.wrapper}>


        <div className={styles.phases} ref={phasesRef}>
          {vertical.map(p => (
            <section key={p.id} className={styles.phase} aria-label={tidy(p.title)}>
              <h2 className={styles.phaseHeader} style={{ margin: 0 }}>
                {isEditMode ? renderEditableTitle(p.id, 'phase', p.title, '(no title)') : tidy(p.title)}
              </h2>
              <div className={styles.columns} style={{ ['--cols' as string]: p.sections.length } as React.CSSProperties}>
                {p.sections.map(sec => renderColumn(sec, p.id))}
              </div>
              {isEditMode && (
                <button type="button" className={styles.addSectionBtn} onClick={() => addSection(p.id)}>+ Add section</button>
              )}
            </section>
          ))}
        </div>

        <div className={styles.footRow}>
          {control.map(p => {
            const sec = p.sections[0];
            const codes = sec ? buildCodeMap(sec) : new Map<string, string>();
            return (
              <section key={p.id} className={`${styles.phase} ${styles.controlPhase}`} aria-label={tidy(p.title)}>
                <h2 className={styles.controlTitle}>
                  {isEditMode ? renderEditableTitle(p.id, 'phase', p.title, '(no title)') : tidy(p.title)}
                </h2>
                <div className={styles.controlInner} style={l3Width ? { ['--l3w' as string]: `${l3Width}px` } as React.CSSProperties : undefined}>
                  {sec && sec.steps.map(s => renderStep(s, codes, { phaseId: p.id, sectionId: sec.id }))}
                  {sec && isEditMode && renderAddStep(p.id, sec.id)}
                </div>
              </section>
            );
          })}
          <img
            className={styles.footLogo}
            src={theme === 'light' ? logoOnLight : logoOnDark}
            alt="Zinfra - Your trusted partner in energy"
          />
        </div>
      </div>

      {isEditMode && (
        <p className={styles.editHint}>
          Click any title or label to rename it. Use × to remove. Step numbers are generated automatically.
        </p>
      )}

      {selected && !isEditMode && (
        <div className={styles.layer}>
          <div className={styles.scrim} onClick={close} />
          <aside
            className={styles.drawer}
            style={{ top: drawerBox.stickyTop, height: `calc(100vh - ${drawerBox.visibleTop}px)` }}
            role="dialog"
            aria-modal="true"
            aria-label="Process details"
          >
            <div className={styles.drawerBar}>
              <button ref={closeRef} type="button" className={styles.closeBtn} aria-label="Close process details" onClick={close}>×</button>
            </div>
            <div className={styles.drawerBody}>
              <ProcessPanelContainer spHttpClient={spHttpClient} webUrl={webUrl} code={selected.code} config={panelConfig} />
            </div>
          </aside>
        </div>
      )}
    </div>
  );
};

export default ProcessHierarchy;
