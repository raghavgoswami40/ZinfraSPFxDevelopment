import * as React from 'react';
import styles from './VerticalFlow.module.scss';
import { IVerticalFlowProps, StoredSection, StoredStep } from './IVerticalFlowProps';
import { FlowButton } from './FlowButton';

const BOX_H = 70;
const GAP   = 8;

const uid = (prefix: string): string => `${prefix}_${Date.now()}_${Math.floor(Math.random() * 9999)}`;

// Split steps into numRows chunks, moving from the end into later rows.
// e.g. 5 steps, 2 rows → [3, 2]; 5 steps, 3 rows → [2, 2, 1]
const splitIntoRows = (steps: StoredStep[], numRows: number): StoredStep[][] => {
  if (steps.length === 0) return [[]];
  const perRow = Math.ceil(steps.length / numRows);
  const rows: StoredStep[][] = [];
  for (let i = 0; i < numRows; i++) {
    const chunk = steps.slice(i * perRow, (i + 1) * perRow);
    if (chunk.length > 0) rows.push(chunk);
  }
  return rows;
};

type EditTarget = { id: string; type: 'step' | 'section' | 'phase' };

const VerticalFlow: React.FC<IVerticalFlowProps> = ({ urls, phases, isEditMode, onPhasesChange }) => {
  const [editing, setEditing]     = React.useState<EditTarget | null>(null);
  const [editValue, setEditValue] = React.useState('');

  // Measured box width — computed from the first section's content area so that
  // exactly 5 boxes fit in one row across every section.
  const [boxW, setBoxW] = React.useState(100);
  const measureRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const el = measureRef.current;
    if (!el) return;
    const compute = (): void => {
      const w = el.clientWidth;
      if (w > 0) {
        // w is the inner content width of sectionContent; 5 boxes + 4 gaps fill it exactly
        setBoxW(Math.max(60, Math.floor((w - 4 * GAP) / 5)));
      }
    };
    compute();
    if (typeof ResizeObserver !== 'undefined') {
      const ro = new ResizeObserver(compute);
      ro.observe(el);
      return () => ro.disconnect();
    }
  }, []);

  // Box style — all steps use the same measured width
  const getStepStyle = (isApproval: boolean): React.CSSProperties => ({
    width: boxW, minWidth: boxW, maxWidth: boxW,
    height: BOX_H, minHeight: BOX_H, maxHeight: BOX_H,
    position: 'relative',
    overflow: 'hidden',
    flexShrink: 0,
    ...(isApproval ? { clipPath: 'polygon(20% 0%, 80% 0%, 100% 50%, 80% 100%, 20% 100%, 0% 50%)' } : {}),
  });

  // Grid row style — fixed boxW columns; sectionContent scrolls if row overflows
  const rowGridStyle = (stepsInRow: number): React.CSSProperties => ({
    display: 'grid',
    gridTemplateColumns: `repeat(${stepsInRow}, ${boxW}px)`,
    gridAutoRows: `${BOX_H}px`,
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
    const updated = phases.map(ph => {
      if (type === 'phase' && ph.id === id) return { ...ph, title: val };
      return {
        ...ph,
        sections: ph.sections.map(sec => {
          if (type === 'section' && sec.id === id) return { ...sec, title: val };
          return {
            ...sec,
            steps: sec.steps.map(st => st.id === id && type === 'step' ? { ...st, label: val } : st),
          };
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
        sections: [...ph.sections, { id, title: 'New\nSection', numRows: 1, steps: [{ id: uid(`${id}_step`), label: 'New step' }] }],
      }
    ));
  };

  const removeSection = (phaseId: string, sectionId: string): void =>
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : { ...ph, sections: ph.sections.filter(s => s.id !== sectionId) }
    ));

  const addRow = (phaseId: string, sectionId: string): void =>
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : {
        ...ph,
        sections: ph.sections.map(sec =>
          sec.id !== sectionId ? sec : { ...sec, numRows: (sec.numRows || 1) + 1 }
        ),
      }
    ));

  const removeRow = (phaseId: string, sectionId: string): void =>
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : {
        ...ph,
        sections: ph.sections.map(sec =>
          sec.id !== sectionId ? sec : { ...sec, numRows: Math.max(1, (sec.numRows || 1) - 1) }
        ),
      }
    ));

  // ── Step actions ──────────────────────────────────────────────────────
  const addStep = (phaseId: string, sectionId: string): void =>
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : {
        ...ph,
        sections: ph.sections.map(sec =>
          sec.id !== sectionId ? sec : { ...sec, steps: [...sec.steps, { id: uid(`${sectionId}_step`), label: 'New step' }] }
        ),
      }
    ));

  const removeStep = (phaseId: string, sectionId: string, stepId: string): void =>
    onPhasesChange(phases.map(ph =>
      ph.id !== phaseId ? ph : {
        ...ph,
        sections: ph.sections.map(sec =>
          sec.id !== sectionId ? sec : { ...sec, steps: sec.steps.filter(st => st.id !== stepId) }
        ),
      }
    ));

  // ── Render step ───────────────────────────────────────────────────────
  const renderStep = (step: StoredStep, phaseId: string, sectionId: string): JSX.Element => {
    const url      = urls[step.id];
    const hasUrl   = !!url;
    const isApproval = !!step.isApproval;
    const isBeingEdited = editing?.id === step.id && editing?.type === 'step';

    const cls = [
      styles.step,
      isApproval ? styles.approval : '',
      !hasUrl && !isEditMode ? styles.noUrl : '',
    ].filter(Boolean).join(' ');

    if (isEditMode) {
      return (
        <div key={step.id} className={cls} style={getStepStyle(isApproval)}>
          {isBeingEdited ? (
            <textarea
              className={styles.stepEditor}
              value={editValue}
              autoFocus
              onChange={e => setEditValue(e.target.value)}
              onBlur={commitEdit}
              onKeyDown={e => {
                if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitEdit(); }
                if (e.key === 'Escape') cancelEdit();
              }}
            />
          ) : (
            <span
              className={styles.stepLabel}
              style={{ cursor: 'pointer' }}
              onClick={() => startEdit(step.id, 'step', step.label)}
              title="Click to edit label"
            >{step.label}</span>
          )}
          <button className={styles.removeBtn} title="Remove step" onClick={() => removeStep(phaseId, sectionId, step.id)}>×</button>
        </div>
      );
    }

    return (
      <div
        key={step.id}
        className={cls}
        style={getStepStyle(isApproval)}
        onClick={() => { if (hasUrl) window.open(url, '_blank', 'noopener noreferrer'); }}
        role={hasUrl ? 'link' : 'presentation'}
        tabIndex={hasUrl ? 0 : -1}
        onKeyDown={e => { if ((e.key === 'Enter' || e.key === ' ') && hasUrl) window.open(url, '_blank', 'noopener noreferrer'); }}
        aria-label={`${step.label}${hasUrl ? ' — open page' : ''}`}
        title={hasUrl ? url : undefined}
      >
        <span className={styles.stepLabel}>{step.label}</span>
      </div>
    );
  };

  // ── Render section ────────────────────────────────────────────────────
  const renderSection = (section: StoredSection, phaseId: string, isFirstSection: boolean): JSX.Element => {
    const goToFlowId  = `${section.id}_goToFlow`;
    const goToFlowUrl = urls[goToFlowId];
    const hasGoToFlow = !!goToFlowUrl;
    const isEditingLabel = editing?.id === section.id && editing?.type === 'section';
    const numRows = section.numRows || 1;
    const stepRows = splitIntoRows(section.steps, numRows);

    return (
      <div key={section.id} className={styles.sectionRow}>

        {/* LEFT: section title + Go to process */}
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

          <FlowButton
            text="Go to process"
            disabled={!hasGoToFlow}
            onClick={() => window.open(goToFlowUrl, '_blank', 'noopener noreferrer')}
          />
        </div>

        {/* RIGHT: steps */}
        <div className={styles.sectionContent}>
          {/* Invisible measure div in the first section — used to compute boxW */}
          {isFirstSection && (
            <div ref={measureRef} style={{ width: '100%', height: 0, visibility: 'hidden' }} />
          )}

          {/* Row controls + add step (edit mode only) */}
          {isEditMode && (
            <div className={styles.sectionActions}>
              <button
                className={styles.rowBtn}
                onClick={() => addRow(phaseId, section.id)}
                title="Add a row (moves last steps down)"
              >+ Row</button>
              <button
                className={`${styles.rowBtn} ${numRows <= 1 ? styles.rowBtnDisabled : ''}`}
                onClick={() => { if (numRows > 1) removeRow(phaseId, section.id); }}
                title={numRows <= 1 ? 'Already at 1 row' : 'Remove last row'}
              >− Row</button>
              <span className={styles.rowCount}>{numRows} row{numRows > 1 ? 's' : ''}</span>
              <button className={styles.addStepBtn} onClick={() => addStep(phaseId, section.id)}>+ Add step</button>
            </div>
          )}

          {/* Steps rendered row by row */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: GAP }}>
            {stepRows.map((rowSteps, rowIdx) => (
              <div key={rowIdx} style={rowGridStyle(rowSteps.length)}>
                {rowSteps.map(step => renderStep(step, phaseId, section.id))}
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
          Click any title to rename · × to remove · use the property pane to set step URLs
        </p>
      )}
    </div>
  );
};

export default VerticalFlow;
