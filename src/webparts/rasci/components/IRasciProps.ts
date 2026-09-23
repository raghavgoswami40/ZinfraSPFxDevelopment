/* eslint-disable @rushstack/no-new-null --
 * assignment is null for "this process has no row in the RASCI list", which is
 * a renderable state distinct from undefined ("not fetched yet"). */

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
 * Why the panel has nothing to show. Drives which unselected message appears —
 * requirement: an L3 with sub-processes must ask for an L4, not for a step.
 */
export type EmptyReason =
  /** Nothing selected in the hierarchy at all. */
  | 'noSelection'
  /** An L3 with L4 children is selected, but no L4 has been opened yet. */
  | 'awaitingL4'
  /** The focused code simply has no row in the RASCI list. */
  | 'notFound';

export interface IRasciProps {
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
  headerIcon?: string;
  colors: RoleColors;

  // ── Unselected states ───────────────────────────────────────────────────
  /** Shown for 'noSelection'. */
  emptyIcon?: string;
  emptyHeadingL3?: string;
  emptyBodyL3?: string;
  /** Shown for 'awaitingL4'. Points up at the process details panel above,
   *  where the L4 step actually gets selected, rather than left at the
   *  process hierarchy — the L3 has already been picked at that point. */
  emptyIconL4?: string;
  emptyHeadingL4?: string;
  emptyBodyL4?: string;
}
