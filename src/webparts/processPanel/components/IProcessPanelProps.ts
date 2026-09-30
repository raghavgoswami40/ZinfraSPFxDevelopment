/* eslint-disable @rushstack/no-new-null --
 * selectedStep/record distinguish "no process selected" (null) from a step
 * that is still loading, and openSubStepCode distinguishes "close every row"
 * (null) from "uncontrolled" (undefined). All null by contract. */

/**
 * Data contracts for the two sections of the Process Panel.
 *
 * Both section components are presentational only: they render whatever data
 * they are handed. Resolving that data (list lookups, logo/system matching) is
 * done by ProcessDetailsService/RasciService before these props are built.
 */

// ── Process Details section ────────────────────────────────────────────────

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
  breadcrumb: string;  // e.g. "Plan › Establish"
  purpose: string;
  inputs: string[];
  outputs: IOutputItem[];   // used only when subSteps is empty
  subSteps: ISubStep[];     // empty array when the step has none
}

export interface IProcessDetailsSectionProps {
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

// ── RASCI section ───────────────────────────────────────────────────────────

/** The five RASCI roles, in the order they are displayed. */
export type RoleKey = 'responsible' | 'accountable' | 'supports' | 'consulted' | 'informed';

export const ROLE_ORDER: RoleKey[] = [
  'responsible', 'accountable', 'supports', 'consulted', 'informed',
];

/** Display label and source column name for each role. The column names are
 *  the ones the RASCI list actually carries, so the two never drift apart. */
export const ROLE_META: Record<RoleKey, { label: string; column: string }> = {
  responsible: { label: 'Responsible', column: 'Responsible' },
  accountable: { label: 'Accountable', column: 'Accountable' },
  supports:    { label: 'Supports',    column: 'Supports' },
  consulted:   { label: 'Consulted',   column: 'Consulted' },
  informed:    { label: 'Informed',    column: 'Informed' },
};

/** One process's RASCI, already split into individual role holders. */
export interface IRasciRecord {
  code: string;         // e.g. "2.1.3.1"
  title: string;        // e.g. "Define Project"
  level: number;        // 3 or 4; 0 when the list leaves Level blank
  /** Role -> the people/teams named for it. Empty array when the cell is blank. */
  roles: Record<RoleKey, string[]>;
}

/** Tag colour per role, as authored in the property pane. */
export type RoleColors = Record<RoleKey, string>;

/**
 * Why the RASCI section has nothing to show. Drives which unselected message
 * appears — requirement: an L3 with sub-processes must ask for an L4, not for
 * a step.
 */
export type EmptyReason =
  /** Nothing selected in the hierarchy at all. */
  | 'noSelection'
  /** An L3 with L4 children is selected, but no L4 has been opened yet. */
  | 'awaitingL4'
  /** The focused code simply has no row in the RASCI list. */
  | 'notFound';

export interface IRasciSectionProps {
  record: IRasciRecord | null;
  isLoading?: boolean;
  error?: string;

  /** Which "nothing to show" message to render. Ignored when record is set. */
  emptyReason: EmptyReason;
  /** The code that produced a 'notFound', for the message. */
  notFoundCode?: string;
  /** Retry handler for the error state. The button is hidden when omitted. */
  onRetry?: () => void;

  // ── Property-pane driven presentation ────────────────────────────────────
  title?: string;
  colors: RoleColors;

  // ── Unselected states ───────────────────────────────────────────────────
  /** Shown for 'noSelection'. */
  emptyIcon?: string;
  emptyHeadingL3?: string;
  emptyBodyL3?: string;
  /** Shown for 'awaitingL4'. Points up at the process details section above,
   *  where the L4 step actually gets selected, rather than left at the
   *  process hierarchy — the L3 has already been picked at that point. */
  emptyIconL4?: string;
  emptyHeadingL4?: string;
  emptyBodyL4?: string;
}
