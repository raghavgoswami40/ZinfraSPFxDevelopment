import * as React from 'react';
import styles from './VerticalFlow.module.scss';
import { IVerticalFlowProps, StoredSection, StoredStep } from './IVerticalFlowProps';

const BOX_H      = 70;
const APPROVAL_H = 46; // Client Approval pills sit shorter than the other shapes
const GAP        = 8;
const ROW_GAP    = 22; // between wrapped rows — leaves room for the code badges above each step
const MAX_COLS   = 5;  // column units per row; boxW is measured so 5 fit exactly

const uid = (prefix: string): string => `${prefix}_${Date.now()}_${Math.floor(Math.random() * 9999)}`;

type StepShape = NonNullable<StoredStep['shape']>;

const getShape = (step: StoredStep): StepShape => step.shape || 'box';

// Per-shape metadata used by the add-step picker and the aria-label suffix.
const SHAPE_META: Record<StepShape, { menuLabel: string; defaultLabel: string; previewClass: string; announce: string }> = {
  box:       { menuLabel: 'Box',              defaultLabel: 'New step',        previewClass: 'previewBox',       announce: '' },
  approval:  { menuLabel: 'Client Approval',  defaultLabel: 'Client Approval', previewClass: 'previewApproval',  announce: ' — Approval step' },
  sapStatus: { menuLabel: 'SAP Status',       defaultLabel: 'STAT',            previewClass: 'previewSapStatus', announce: ' — SAP status step' },
  gate:      { menuLabel: 'Gate',             defaultLabel: 'Gate',            previewClass: 'previewGate',      announce: ' — Gate step' },
  group:     { menuLabel: 'Dashed group',     defaultLabel: 'Group title',     previewClass: 'previewGroup',     announce: ' — Group' },
};

// How many grid columns each shape occupies. The dashed group holds nested
// steps side by side, so it needs the width of two normal boxes.
const COL_SPAN: Record<StepShape, number> = { box: 1, approval: 1, sapStatus: 1, gate: 1, group: 2 };

const getSpan = (step: StoredStep): number => COL_SPAN[getShape(step)];

// Pack steps greedily: fill each row up to MAX_COLS column units before
// wrapping, so earlier rows are full rather than every row being equal length.
const packRows = (steps: StoredStep[]): StoredStep[][] => {
  const rows: StoredStep[][] = [];
  let row: StoredStep[] = [];
  let used = 0;
  steps.forEach(step => {
    const span = getSpan(step);
    if (row.length > 0 && used + span > MAX_COLS) {
      rows.push(row);
      row = [];
      used = 0;
    }
    row.push(step);
    used += span;
  });
  if (row.length > 0) rows.push(row);
  return rows.length > 0 ? rows : [[]];
};

// A section numbers its steps "<prefix>.1", "<prefix>.2", ... The prefix is
// taken from codePrefix, or derived from a title that starts with the number.
const getCodePrefix = (section: StoredSection): string | undefined => {
  if (section.codePrefix) return section.codePrefix;
  const first = (section.title || '').trim().split(/[\s\n]+/)[0];
  return /^\d+(\.\d+)*$/.test(first) ? first : undefined;
};

// Codes run in display order across every step that carries a badge. Gate
// diamonds are never numbered, and a dashed group passes its number on to the
// steps nested inside it rather than taking one itself.
const buildCodeMap = (section: StoredSection): Map<string, string> => {
  const map = new Map<string, string>();
  const prefix = getCodePrefix(section);
  if (!prefix) return map;
  let n = 0;
  const visit = (steps: StoredStep[]): void => {
    steps.forEach(step => {
      const shape = getShape(step);
      if (shape === 'gate') return;
      if (shape === 'group') { visit(step.children || []); return; }
      n += 1;
      map.set(step.id, `${prefix}.${n}`);
    });
  };
  visit(section.steps);
  return map;
};

type EditTarget = { id: string; type: 'step' | 'section' | 'phase' };

const VerticalFlow: React.FC<IVerticalFlowProps> = ({ phases, isEditMode, onPhasesChange }) => {
  const [editing, setEditing]     = React.useState<EditTarget | null>(null);
  const [editValue, setEditValue] = React.useState('');

  // Measured box width — computed from the first section's content area so that
  // exactly 5 boxes fit in one row across every section.
  const [boxW, setBoxW] = React.useState(100);
  const measureRef = React.useRef<HTMLDivElement>(null);

  // Which section's "+ Add step" shape-picker menu is currently open (if any)
  const [addMenuSectionId, setAddMenuSectionId] = React.useState<string | null>(null);
  const addMenuRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (!addMenuSectionId) return;
    const onDocMouseDown = (e: MouseEvent): void => {
      if (addMenuRef.current && !addMenuRef.current.contains(e.target as Node)) {
        setAddMenuSectionId(null);
      }
    };
    document.addEventListener('mousedown', onDocMouseDown);
    return () => document.removeEventListener('mousedown', onDocMouseDown);
  }, [addMenuSectionId]);

  React.useEffect(() => {
    const el = measureRef.current;
    if (!el) return;
    const compute = (): void => {
      const w = el.clientWidth;
      if (w > 0) {
        // w is the inner content width of sectionContent; 5 boxes + 4 gaps fill it exactly
        setBoxW(Math.max(60, Math.floor((w - (MAX_COLS - 1) * GAP) / MAX_COLS)));
      }
    };
    compute();
    if (typeof ResizeObserver !== 'undefined') {
      const ro = new ResizeObserver(compute);
      ro.observe(el);
      return () => ro.disconnect();
    }
  }, []);

  // Step sizing. Shape geometry (pill/hexagon/diamond) is drawn by a clipped
  // ::before background layer in CSS, so these elements stay unclipped and can
  // carry badges that sit outside the shape's own outline.
  const getStepStyle = (shape: StepShape): React.CSSProperties => {
    // The dashed group spans two columns and is taller than one row (title plus
    // full-size nested steps), so let its content define the height.
    if (shape === 'group') {
      return {
        gridColumn: 'span 2',
        position: 'relative',
        overflow: 'visible',
        alignSelf: 'center',
      };
    }
    const h = shape === 'approval' ? APPROVAL_H : BOX_H;
    return {
      width: boxW, minWidth: boxW, maxWidth: boxW,
      height: h, minHeight: h, maxHeight: h,
      position: 'relative',
      overflow: 'visible',
      alignSelf: 'center',
      flexShrink: 0,
    };
  };

  // Grid row style — column count is the sum of the row's shape spans, so a
  // 2-wide group doesn't push the row into an implicit overflowing column.
  // Rows grow past BOX_H when they contain a taller item (i.e. a dashed group).
  const rowGridStyle = (rowSteps: StoredStep[]): React.CSSProperties => ({
    display: 'grid',
    gridTemplateColumns: `repeat(${rowSteps.reduce((sum, st) => sum + getSpan(st), 0)}, ${boxW}px)`,
    gridAutoRows: `minmax(${BOX_H}px, auto)`,
    alignItems: 'center',
    gap: GAP,
  });

  const startEdit = (id: string, type: EditTarget['type'], current: string): void => {
    setEditing({ id, type });
    setEditValue(current);
  };

  const commitEdit = (): void => {
    if (!editing) return;
    const { id, type } = editing;
    const val = editValue.trim() || '(untitled)';
    // Recurses so steps nested inside a dashed group are editable too.
    const editStep = (st: StoredStep): StoredStep => {
      let next = st.id === id && type === 'step' ? { ...st, label: val } : st;
      if (next.children) next = { ...next, children: next.children.map(editStep) };
      return next;
    };
    const updated = phases.map(ph => {
      if (type === 'phase' && ph.id === id) return { ...ph, title: val };
      return {
        ...ph,
        sections: ph.sections.map(sec => {
          if (type === 'section' && sec.id === id) return { ...sec, title: val };
          return { ...sec, steps: sec.steps.map(editStep) };
        }),
      };
    });
    onPhasesChange(updated);
    setEditing(null);
  };

  const cancelEdit = (): void => setEditing(null);

  // ── Section actions ───────────────────────────────────────────────────
  const addSection = (phaseId: string): void => {
    const id = uid(`${phaseId}_sec`);
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : {
        ...ph,
        sections: [...ph.sections, { id, title: 'New\nSection', steps: [{ id: uid(`${id}_step`), label: 'New step' }] }],
      }
    ));
  };

  const removeSection = (phaseId: string, sectionId: string): void =>
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : { ...ph, sections: ph.sections.filter(s => s.id !== sectionId) }
    ));

  // ── Step actions ──────────────────────────────────────────────────────
  const addStep = (phaseId: string, sectionId: string, shape: StepShape = 'box'): void => {
    const id = uid(`${sectionId}_step`);
    const newStep: StoredStep = { id, label: SHAPE_META[shape].defaultLabel, shape };
    if (shape === 'group') {
      newStep.children = [
        { id: `${id}_c1`, label: 'STAT', shape: 'sapStatus' },
        { id: `${id}_c2`, label: 'STAT', shape: 'sapStatus' },
      ];
    }
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : {
        ...ph,
        sections: ph.sections.map(sec =>
          sec.id !== sectionId ? sec : { ...sec, steps: [...sec.steps, newStep] }
        ),
      }
    ));
    setAddMenuSectionId(null);
  };

  // Recurses so a step nested inside a dashed group can be removed too.
  const filterSteps = (steps: StoredStep[], stepId: string): StoredStep[] =>
    steps
      .filter(st => st.id !== stepId)
      .map(st => st.children ? { ...st, children: filterSteps(st.children, stepId) } : st);

  const removeStep = (phaseId: string, sectionId: string, stepId: string): void =>
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : {
        ...ph,
        sections: ph.sections.map(sec =>
          sec.id !== sectionId ? sec : { ...sec, steps: filterSteps(sec.steps, stepId) }
        ),
      }
    ));

  // A single editable text node — used for step labels, group titles, and the
  // status text of steps nested inside a dashed group.
  const renderEditableText = (step: StoredStep, className: string, hint: string, editorClassName: string = styles.stepEditor): JSX.Element => {
    if (isEditMode && editing?.id === step.id && editing?.type === 'step') {
      return (
        <textarea
          className={editorClassName}
          value={editValue}
          autoFocus
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
      <span
        className={className}
        style={isEditMode ? { cursor: 'pointer' } : undefined}
        onClick={isEditMode ? () => startEdit(step.id, 'step', step.label) : undefined}
        title={isEditMode ? hint : undefined}
      >{step.label}</span>
    );
  };

  // Inner content of a step. SAP status steps carry a fixed "SAP Status" first
  // line above the editable status text; `noPrefix` opts a step out of that
  // (used where the hexagon names something other than a SAP status).
  const renderStepContent = (step: StoredStep, hint: string): JSX.Element => {
    const showSapPrefix = getShape(step) === 'sapStatus' && !step.noPrefix;
    return (
      <>
        {showSapPrefix && <span className={styles.sapPrefix}>SAP Status</span>}
        {renderEditableText(step, styles.stepLabel, hint)}
      </>
    );
  };

  // ── Render dashed group ───────────────────────────────────────────────
  // A dashed container with its own title and full-size nested steps inside it.
  const renderGroup = (step: StoredStep, phaseId: string, sectionId: string, codes: Map<string, string>): JSX.Element => (
    <div key={step.id} className={styles.shapeGroup} style={getStepStyle('group')}>
      {renderEditableText(step, styles.groupTitle, 'Click to edit group title', styles.groupTitleEditor)}
      <div className={styles.groupChildren}>
        {(step.children || []).map(child => (
          <div
            key={child.id}
            className={`${styles.step} ${styles.shapeSapStatus}`}
            style={{ flex: 1, minWidth: 0, height: BOX_H, position: 'relative', overflow: 'visible' }}
          >
            {codes.get(child.id) && <span className={styles.stepCodeBadge}>{codes.get(child.id)}</span>}
            {renderStepContent(child, 'Click to edit status')}
            {isEditMode && (
              <button
                className={styles.removeBtnSmall}
                title="Remove from group"
                onClick={() => removeStep(phaseId, sectionId, child.id)}
              >×</button>
            )}
          </div>
        ))}
      </div>
      {isEditMode && (
        <button className={styles.removeBtn} title="Remove group" onClick={() => removeStep(phaseId, sectionId, step.id)}>×</button>
      )}
    </div>
  );

  // ── Render step ───────────────────────────────────────────────────────
  const renderStep = (step: StoredStep, phaseId: string, sectionId: string, codes: Map<string, string>): JSX.Element => {
    const shape = getShape(step);
    if (shape === 'group') return renderGroup(step, phaseId, sectionId, codes);

    const code = codes.get(step.id);

    const cls = [
      styles.step,
      shape === 'approval'  ? styles.shapeApproval  : '',
      shape === 'sapStatus' ? styles.shapeSapStatus : '',
      shape === 'gate'      ? styles.shapeGate      : '',
    ].filter(Boolean).join(' ');

    const codeBadge = code ? <span className={styles.stepCodeBadge}>{code}</span> : null;

    if (isEditMode) {
      return (
        <div key={step.id} className={cls} style={getStepStyle(shape)}>
          {codeBadge}
          {renderStepContent(step, 'Click to edit label')}
          <button className={styles.removeBtn} title="Remove step" onClick={() => removeStep(phaseId, sectionId, step.id)}>×</button>
        </div>
      );
    }

    const sapPrefix = shape === 'sapStatus' && !step.noPrefix;
    return (
      <div
        key={step.id}
        className={cls}
        style={getStepStyle(shape)}
        aria-label={`${code ? code + ' — ' : ''}${step.label}${SHAPE_META[shape].announce}`}
      >
        {codeBadge}
        {sapPrefix && <span className={styles.sapPrefix}>SAP Status</span>}
        <span className={styles.stepLabel}>{step.label}</span>
      </div>
    );
  };

  // ── Render section ────────────────────────────────────────────────────
  const renderSection = (section: StoredSection, phaseId: string, isFirstSection: boolean): JSX.Element => {
    const isEditingLabel = editing?.id === section.id && editing?.type === 'section';
    const codes = buildCodeMap(section);
    const stepRows = packRows(section.steps);
    const sectionCls = [
      styles.sectionRow,
      section.variant === 'grey' ? styles.sectionRowGrey : '',
    ].filter(Boolean).join(' ');

    return (
      <div key={section.id} className={sectionCls}>

        {/* LEFT: section title */}
        <div className={styles.sectionLabel}>
          {isEditMode && isEditingLabel ? (
            <textarea
              className={styles.sectionLabelEditing}
              value={editValue}
              autoFocus
              rows={3}
              onChange={e => setEditValue(e.target.value)}
              onBlur={commitEdit}
              onKeyDown={e => {
                if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitEdit(); }
                if (e.key === 'Escape') cancelEdit();
              }}
            />
          ) : (
            <span
              style={isEditMode ? { cursor: 'pointer' } : {}}
              onClick={isEditMode ? () => startEdit(section.id, 'section', section.title) : undefined}
              title={isEditMode ? 'Click to rename' : undefined}
            >{section.title}</span>
          )}
        </div>

        {/* RIGHT: steps */}
        <div className={styles.sectionContent}>
          {/* Invisible measure div in the first section — used to compute boxW */}
          {isFirstSection && (
            <div ref={measureRef} style={{ width: '100%', height: 0, visibility: 'hidden' }} />
          )}

          {/* Add step (edit mode only) */}
          {isEditMode && (
            <div className={styles.sectionActions}>
              <div
                className={styles.addStepWrap}
                ref={addMenuSectionId === section.id ? addMenuRef : undefined}
              >
                <button
                  className={styles.addStepBtn}
                  onClick={() => setAddMenuSectionId(addMenuSectionId === section.id ? null : section.id)}
                >+ Add step</button>
                {addMenuSectionId === section.id && (
                  <div className={styles.addMenu}>
                    {(Object.keys(SHAPE_META) as StepShape[]).map(shape => (
                      <button
                        key={shape}
                        className={styles.addMenuItem}
                        onClick={() => addStep(phaseId, section.id, shape)}
                      >
                        <span className={`${styles.addMenuShapePreview} ${(styles as Record<string, string>)[SHAPE_META[shape].previewClass]}`} />
                        {SHAPE_META[shape].menuLabel}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Steps rendered row by row */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: ROW_GAP }}>
            {stepRows.map((rowSteps, rowIdx) => (
              <div key={rowIdx} style={rowGridStyle(rowSteps)}>
                {rowSteps.map(step => renderStep(step, phaseId, section.id, codes))}
              </div>
            ))}
          </div>
        </div>

        {/* Remove section (edit mode only) */}
        {isEditMode && (
          <div className={styles.removeSectionBtnWrap}>
            <button className={styles.iconBtn} title="Remove section" onClick={() => removeSection(phaseId, section.id)}>×</button>
          </div>
        )}
      </div>
    );
  };

  // Track the globally first section for the measure ref
  const firstSectionId = phases[0]?.sections[0]?.id;

  return (
    <div className={styles.wrapper}>
      {phases.map(phase => {
        const isEditingTitle = editing?.id === phase.id && editing?.type === 'phase';
        return (
          <div key={phase.id} className={styles.phase}>
            <div className={styles.phaseHeader}>
              {isEditMode && isEditingTitle ? (
                <input
                  className={styles.phaseTitleEditing}
                  value={editValue}
                  autoFocus
                  onChange={e => setEditValue(e.target.value)}
                  onBlur={commitEdit}
                  onKeyDown={e => { if (e.key === 'Enter') commitEdit(); if (e.key === 'Escape') cancelEdit(); }}
                />
              ) : (
                <span
                  style={isEditMode ? { cursor: 'pointer' } : {}}
                  onClick={isEditMode ? () => startEdit(phase.id, 'phase', phase.title) : undefined}
                  title={isEditMode ? 'Click to rename' : undefined}
                >{phase.title}</span>
              )}
            </div>

            <div className={styles.phaseSections}>
              {phase.sections.map(sec =>
                renderSection(sec, phase.id, sec.id === firstSectionId)
              )}
              {isEditMode && (
                <button className={styles.addSectionBtn} onClick={() => addSection(phase.id)}>+ Add section</button>
              )}
            </div>
          </div>
        );
      })}
      {isEditMode && (
        <p className={styles.editHint}>
          Click any title to rename · × to remove · step numbers are generated automatically
        </p>
      )}
    </div>
  );
};

export default VerticalFlow;
