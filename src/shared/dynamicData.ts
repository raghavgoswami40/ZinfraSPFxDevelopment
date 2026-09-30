/**
 * Contract between the Vertical Flow web part (Dynamic Data source) and the
 * Process Panel web part (consumer).
 *
 * Both bundles import this module, so the property id and component id exist in
 * exactly one place. A mismatch between them produces an empty sources dropdown
 * with no error message at all, which is close to undiagnosable.
 *
 * There used to be a second contract here, between Process Details and RASCI —
 * they were separate web parts, so the L3/L4 focus one computed had to be
 * published as Dynamic Data for the other to read. Process Panel merged them
 * into a single instance, so that hand-off is now just internal state; see
 * ProcessPanelWebPart's _resolveRasciTarget().
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
