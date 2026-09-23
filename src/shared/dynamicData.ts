/**
 * Contract shared between the Vertical Flow web part (Dynamic Data source) and
 * the Process Details web part (consumer).
 *
 * Both bundles import this module, so the property id and component id exist in
 * exactly one place. A mismatch between them produces an empty sources dropdown
 * with no error message at all, which is close to undiagnosable.
 */

/** Id of the property Vertical Flow publishes. Letters, digits, - and _ only. */
export const SELECTED_STEP_PROPERTY_ID: string = 'selectedProcessStep';

/**
 * Manifest id of VerticalFlowWebPart, used to filter the property pane's source
 * dropdown. Must match VerticalFlowWebPart.manifest.json "id" character for
 * character — lowercase, no braces.
 */
export const VERTICAL_FLOW_COMPONENT_ID: string = 'bbfa2c75-800f-44da-81eb-52bdd4ba6a9c';

/**
 * Payload published on SELECTED_STEP_PROPERTY_ID.
 *
 * Must stay a flat, JSON-serialisable object: the framework deep-clones it on
 * every read, so no Map, no class instances, no functions, no cycles.
 */
export interface ISelectedProcessStep {
  /** StoredStep.id — the stable identity. Codes renumber when the diagram is
   *  edited; this does not, so it is what drives the source's own highlight. */
  stepId: string;
  /** The derived badge code, e.g. "2.1.3" — the key the details panel looks up
   *  in SharePoint. Absent for a step in a section with no resolvable prefix. */
  code?: string;
  /** The step label as displayed. */
  label: string;
  /** Raw section title, e.g. "2.1 ESTABLISH". */
  sectionTitle: string;
  /** Raw phase title, e.g. "2  PLAN". */
  phaseTitle: string;
}

/* ────────────────────────────────────────────────────────────────────────────
 * Second contract: Process Details (source) -> RASCI (consumer).
 *
 * RASCI needs to know two things that only Process Details can answer: which
 * L3 is selected, and — because an L3 with L4 children must show its L4's
 * RASCI rather than its own — whether that L3 has children and which one the
 * reader has opened. Publishing both from Process Details means RASCI needs a
 * single connection instead of two, and can never disagree with the panel
 * about whether sub-steps exist.
 * ──────────────────────────────────────────────────────────────────────────── */

/** Id of the property Process Details publishes. */
export const PROCESS_FOCUS_PROPERTY_ID: string = 'processFocus';

/**
 * Manifest id of ProcessDetailsWebPart, used to filter RASCI's source
 * dropdown. Must match ProcessDetailsWebPart.manifest.json "id" exactly.
 */
export const PROCESS_DETAILS_COMPONENT_ID: string = '7adf9e52-9cf5-4951-adbf-b3bf8ab2dab5';

/**
 * Payload published on PROCESS_FOCUS_PROPERTY_ID. Flat and JSON-serialisable,
 * for the same reason as ISelectedProcessStep.
 */
export interface IProcessFocus {
  /** Code of the L3 selected in the hierarchy; absent when nothing is selected. */
  l3Code?: string;
  l3Title?: string;
  /** True once the L3's row has loaded and turned out to have L4 children.
   *  False while still loading, and for a childless L3. */
  hasSubSteps: boolean;
  /** True while the L3's own details are still being fetched, so a consumer
   *  can hold its empty state rather than flashing the wrong one. */
  pending: boolean;
  /** The L4 row currently expanded in the panel, if any. */
  l4Code?: string;
  l4Title?: string;
  /**
   * The code a consumer should actually display:
   *  - the open L4 when the L3 has children,
   *  - the L3 itself when it has none,
   *  - absent when nothing is selected, or an L3 with children has no open L4.
   */
  focusCode?: string;
  focusTitle?: string;
}
