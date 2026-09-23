/* eslint-disable @rushstack/no-new-null --
 * selectedStep distinguishes "no process selected" (null) from a step that is
 * still loading, and openSubStepCode distinguishes "close every row" (null)
 * from "uncontrolled" (undefined). Both are null by contract. */

/**
 * Data contract for the Process Details panel.
 *
 * The component is presentational only: it renders whatever IProcessStep it is
 * handed. Resolving each output to its artefact/system record (logo + link) is
 * done by ProcessDetailsService before these props are built.
 */

export interface IOutputItem {
  text: string;
  /** System tag from "Output N - System"; "" when the row carries none. */
  system: string;
  /** Logo resolved from the System Master list; undefined when unmatched. */
  logoUrl?: string;
  /** "Output N - Link"; empty when no URL has been recorded. */
  href?: string;
}

export interface ISubStep {
  code: string;        // e.g. "2.1.3.1"
  title: string;
  outputs: IOutputItem[];
}

export interface IProcessStep {
  code: string;        // e.g. "2.1.3"
  title: string;       // e.g. "Define Project"
  breadcrumb: string;  // e.g. "PLAN › Establish"
  purpose: string;
  inputs: string[];
  outputs: IOutputItem[];   // used only when subSteps is empty
  subSteps: ISubStep[];     // empty array when the step has none
}

export interface IProcessDetailsProps {
  selectedStep: IProcessStep | null;
  isLoading?: boolean;
  error?: string;

  /** Set when a code was selected but the list holds no rows for it — a
   *  legitimate "not authored yet" state, distinct from nothing being selected. */
  notFoundCode?: string;

  /** Controlled mode: when supplied (not undefined) this is the source of truth
   *  for which L4 row is expanded. */
  openSubStepCode?: string | null;
  /** Fires with the newly opened code, or null when the row was closed. */
  onSubStepToggle?: (code: string | null) => void;
  /** Retry handler for the error state. The button is hidden when omitted. */
  onRetry?: () => void;

  // ── Property-pane driven presentation ────────────────────────────────────
  title?: string;             // header label, default "Process details"
  headerIcon?: string;        // header emoji/glyph, default "🧾"
  showBreadcrumb?: boolean;   // default true
  singleOpenSubStep?: boolean;// default true — false allows multiple open rows

  // ── Unselected ("nothing picked yet") state, also property-pane driven ───
  /** Emoji/glyph shown above the heading. Default "👈". */
  emptyIcon?: string;
  /** Default "Select a process to view details". */
  emptyHeading?: string;
  /** Default "Click on any process in the process hierarchy to see its
   *  inputs, procedure and outputs here." */
  emptyBody?: string;
}
