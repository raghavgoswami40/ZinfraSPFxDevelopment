import { StoredPhase, StoredSection, StoredStep } from './components/IVerticalFlowProps';
import { ISelectedProcessStep } from '../../shared/dynamicData';

/**
 * Step numbering, extracted from VerticalFlow.tsx so the web part class can
 * derive the same codes the renderer draws on the badges. Two implementations
 * of this numbering would drift, and the drift would be silent — the panel
 * would simply show a different process than the one that looks selected.
 */

export type StepShape = NonNullable<StoredStep['shape']>;

export const getShape = (step: StoredStep): StepShape => step.shape || 'box';

// A section numbers its steps "<prefix>.1", "<prefix>.2", ... The prefix is
// taken from codePrefix, or derived from a title that starts with the number.
export const getCodePrefix = (section: StoredSection): string | undefined => {
  if (section.codePrefix) return section.codePrefix;
  const first = (section.title || '').trim().split(/[\s\n]+/)[0];
  return /^\d+(\.\d+)*$/.test(first) ? first : undefined;
};

// Codes run in display order across every step that carries a badge. Gate
// diamonds are never numbered, and a dashed group passes its number on to the
// steps nested inside it rather than taking one itself.
export const buildCodeMap = (section: StoredSection): Map<string, string> => {
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

// Mirrors filterSteps in VerticalFlow.tsx — a step can be nested inside a
// dashed group, so the search has to recurse.
const findStep = (steps: StoredStep[], id: string): StoredStep | undefined => {
  for (const step of steps) {
    if (step.id === id) return step;
    const child = step.children ? findStep(step.children, id) : undefined;
    if (child) return child;
  }
  return undefined;
};

/**
 * Rebuilds the published payload from the current phases every time, rather
 * than caching it at click time. That is what keeps the payload honest when the
 * diagram is edited underneath a live selection: renaming the step updates the
 * label, and deleting it returns undefined — which is exactly the signal the
 * caller needs to clear the selection.
 */
export const resolveSelectedStep = (
  phases: StoredPhase[],
  stepId: string | undefined
): ISelectedProcessStep | undefined => {
  if (!stepId) return undefined;

  for (const phase of phases) {
    for (const section of phase.sections) {
      const found = findStep(section.steps, stepId);
      if (found) {
        return {
          stepId: found.id,
          code: buildCodeMap(section).get(found.id),
          label: found.label,
          // Collapsed the same way the section banner collapses it, so the
          // payload matches what the reader actually sees on screen.
          sectionTitle: (section.title || '').replace(/\s*\n\s*/g, ' ').trim(),
          phaseTitle: phase.title,
        };
      }
    }
  }
  return undefined;
};
