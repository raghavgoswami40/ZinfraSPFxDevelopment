/**
 * Data contracts for the Master RASCI web part.
 *
 * The SharePoint "RASCI" list holds one row per process with five role columns
 * (who is Responsible, Accountable, ...). The Master RASCI view is the inverse:
 * for every *role*, which code does it hold on each activity. MasterRasciService
 * does that inversion once on load; everything downstream works on these shapes.
 */

/** A role's involvement in one activity. AR counts as Accountable and Responsible. */
export type RasciCode = 'A' | 'AR' | 'R' | 'S' | 'C' | 'I';

/** The five single codes, which are the five columns of the view. */
export type ChipCode = 'A' | 'R' | 'S' | 'C' | 'I';

export interface IRasciRole {
  name: string;
  /** Role group shown as the band above the Matrix columns. */
  group: string;
}

export interface IRasciActivity {
  /** SharePoint item id — stable across loads, so safe as a React key. */
  id: number;
  /** L1, e.g. "2 Plan". */
  phase: string;
  /** L2, e.g. "2.1 Establish". */
  group: string;
  /** L3 process, e.g. "Define project". */
  process: string;
  /** L4 activity; empty when the L3 row is itself the activity. */
  activity: string;
  /** Role index (into IMasterRasciData.roles) -> code. Missing = no involvement. */
  assignments: { [roleIndex: number]: RasciCode };
}

export interface IMasterRasciData {
  roles: IRasciRole[];
  rows: IRasciActivity[];
}

export type RasciView = 'list' | 'matrix';

export interface IFilterState {
  phase: string;          // 'All' or a phase value
  group: string;          // 'All' or a process group value
  role: number;           // -1 for all roles, else an index into roles
  q: string;
}

/** Texts a page author can rename in the property pane. */
export interface IMasterRasciLabels {
  phase: string;
  search: string;
  group: string;
  role: string;
  viewList: string;
  viewMatrix: string;
}

export const DEFAULT_LABELS: IMasterRasciLabels = {
  phase: 'Phase',
  search: 'Search process',
  group: 'Process group',
  role: 'View as role',
  viewList: 'By process',
  viewMatrix: 'Matrix',
};

export interface IMasterRasciProps {
  /** Resolves the data; the component owns loading and error state. */
  load: () => Promise<IMasterRasciData>;
  listTitle: string;
  documentRef: string;
  defaultView: RasciView;
  collapseAfter: number;
  labels: IMasterRasciLabels;
}
