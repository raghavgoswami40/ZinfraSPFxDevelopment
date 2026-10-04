import {
  ChipCode, RasciCode, IMasterRasciData, IRasciActivity, IFilterState,
} from './IMasterRasciProps';

/**
 * Pure filter/derive logic for the Master RASCI view, ported from the design
 * prototype's renderVals(). No React and no DOM, so it can be exercised directly.
 * Written for an ES5 target with a narrow lib: no Object.entries, String.includes
 * or Array.find.
 */

export interface ICodeMeta {
  code: ChipCode;
  name: string;
  /** Badge / chip border colour. */
  color: string;
  /** Badge text colour. */
  fg: string;
  /** Chip background and text. */
  chipBg: string;
  chipFg: string;
}

/** Same colour profile as the Process Panel's RASCI section (its defaults). */
export const CODES: ICodeMeta[] = [
  { code: 'A', name: 'Accountable', color: '#BFBFBF', fg: '#0B1B26', chipBg: '#BFBFBF', chipFg: '#0B1B26' },
  { code: 'R', name: 'Responsible', color: '#016891', fg: '#FFFFFF', chipBg: '#016891', chipFg: '#FFFFFF' },
  { code: 'S', name: 'Support',     color: '#009DDC', fg: '#0B1B26', chipBg: '#009DDC', chipFg: '#0B1B26' },
  { code: 'C', name: 'Consulted',   color: '#6BA244', fg: '#0B1B26', chipBg: '#6BA244', chipFg: '#0B1B26' },
  { code: 'I', name: 'Informed',    color: '#B3BE35', fg: '#0B1B26', chipBg: '#B3BE35', chipFg: '#0B1B26' },
];

export const HIGHLIGHT = '#FFC000';

/** AR is both Accountable and Responsible; every other code matches only itself. */
export const codeMatches = (held: RasciCode, wanted: ChipCode): boolean =>
  held === wanted || (held === 'AR' && (wanted === 'A' || wanted === 'R'));

export const EMPTY_FILTERS: IFilterState = { phase: 'All', group: 'All', role: -1, q: '' };

export const hasActiveFilters = (f: IFilterState): boolean =>
  f.phase !== 'All' || f.group !== 'All' || f.role >= 0 || f.q !== '';

/** (roleIndex, code) pairs for an activity, in role order. */
export const assignmentsOf = (row: IRasciActivity): Array<{ role: number; code: RasciCode }> => {
  const out: Array<{ role: number; code: RasciCode }> = [];
  for (const key in row.assignments) {
    if (Object.prototype.hasOwnProperty.call(row.assignments, key)) {
      out.push({ role: Number(key), code: row.assignments[key] });
    }
  }
  return out.sort((a, b) => a.role - b.role);
};

/**
 * All filters combine with AND.
 */
export const filterRows = (
  data: IMasterRasciData,
  f: IFilterState,
  ignoreGroup: boolean = false
): IRasciActivity[] => {
  const ql = f.q.trim().toLowerCase();
  return data.rows.filter((r) => {
    if (f.phase !== 'All' && r.phase !== f.phase) { return false; }
    if (!ignoreGroup && f.group !== 'All' && r.group !== f.group) { return false; }
    // Case-insensitive match across process group, process and activity.
    if (ql !== '' && (r.group + ' ' + r.process + ' ' + r.activity).toLowerCase().indexOf(ql) === -1) {
      return false;
    }
    if (f.role >= 0 && r.assignments[f.role] === undefined) { return false; }
    return true;
  });
};

export interface IPhaseOption { key: string; label: string; count: number }

/** "All phases" plus one entry per phase, counted over the whole data set — the
 *  counts deliberately ignore the other filters. */
export const phaseOptions = (data: IMasterRasciData): IPhaseOption[] => {
  const order: string[] = [];
  const counts: { [k: string]: number } = {};
  data.rows.forEach((r) => {
    if (counts[r.phase] === undefined) { counts[r.phase] = 0; order.push(r.phase); }
    counts[r.phase]++;
  });
  return [{ key: 'All', label: 'All phases', count: data.rows.length }]
    .concat(order.map((p) => ({ key: p, label: p, count: counts[p] })));
};

export interface IGroupOption { phase: string; groups: string[] }

/** Process groups bucketed by phase, in data order, for the <optgroup> select. */
export const groupOptions = (data: IMasterRasciData): IGroupOption[] => {
  const out: IGroupOption[] = [];
  data.rows.forEach((r) => {
    let bucket: IGroupOption | undefined;
    for (const b of out) { if (b.phase === r.phase) { bucket = b; break; } }
    if (!bucket) { bucket = { phase: r.phase, groups: [] }; out.push(bucket); }
    if (bucket.groups.indexOf(r.group) === -1) { bucket.groups.push(r.group); }
  });
  return out;
};

/** The phase a process group belongs to, or undefined for an unknown group. */
export const phaseOfGroup = (data: IMasterRasciData, group: string): string | undefined => {
  for (const r of data.rows) { if (r.group === group) { return r.phase; } }
  return undefined;
};

/** A heading row followed by its activities, in display order. */
export type GroupedRow =
  | { kind: 'head'; key: string; title: string; count: number }
  | { kind: 'row'; key: string; row: IRasciActivity };

/** Inserts a heading each time the process group changes. */
export const groupRows = (rows: IRasciActivity[]): GroupedRow[] => {
  const counts: { [g: string]: number } = {};
  rows.forEach((r) => { counts[r.group] = (counts[r.group] || 0) + 1; });

  const out: GroupedRow[] = [];
  let last: string | undefined;
  rows.forEach((r) => {
    if (r.group !== last) {
      last = r.group;
      out.push({ kind: 'head', key: 'h:' + r.group + ':' + r.id, title: r.group, count: counts[r.group] });
    }
    out.push({ kind: 'row', key: 'r:' + r.id, row: r });
  });
  return out;
};

/** Roles with at least one assignment in the current result set (Matrix columns). */
export const visibleRoles = (data: IMasterRasciData, rows: IRasciActivity[]): number[] => {
  const used: { [i: number]: boolean } = {};
  rows.forEach((r) => assignmentsOf(r).forEach((a) => { used[a.role] = true; }));
  return data.roles.map((_, i) => i).filter((i) => used[i]);
};

/** Consecutive visible roles that share a group, for the Matrix band row. */
export const roleBands = (data: IMasterRasciData, roles: number[]): Array<{ name: string; span: number }> => {
  const bands: Array<{ name: string; span: number }> = [];
  roles.forEach((i) => {
    const g = data.roles[i].group;
    const last = bands[bands.length - 1];
    if (last && last.name === g) { last.span++; } else { bands.push({ name: g, span: 1 }); }
  });
  return bands;
};

export interface IColumnEntry { role: number; code: RasciCode }

/**
 * Who holds `code` on this activity, the selected role first. An AR role appears
 * under both A and R.
 */
export const columnEntries = (row: IRasciActivity, code: ChipCode, selectedRole: number): IColumnEntry[] =>
  assignmentsOf(row)
    .filter((a) => codeMatches(a.code, code))
    .map((a) => ({ role: a.role, code: a.code }))
    .sort((a, b) => Number(b.role === selectedRole) - Number(a.role === selectedRole));

