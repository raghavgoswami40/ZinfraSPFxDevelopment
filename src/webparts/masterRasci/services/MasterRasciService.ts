import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  ListFieldMap, loadFieldMap, IFieldInfo, isLookupType,
} from '../../../shared/spFieldNames';
import {
  ChipCode, RasciCode, IMasterRasciData, IRasciActivity, IRasciRole,
} from '../components/IMasterRasciProps';

/**
 * Reads the whole "RASCI" list and reshapes it for the Master RASCI view.
 *
 * The list stores one row per process with five role columns, each holding the
 * names of the roles that carry that code (semicolon separated). The Master view
 * needs the opposite: per activity, which code each *role* holds. This service
 * inverts the list once on load; filtering afterwards is entirely client side.
 */

/** Column display names. Only these need to match the list. */
const COL = {
  title: 'Title',
  processCode: 'ProcessCode',
  parentProcessCode: 'ParentProcessCode',
  level: 'Level',
  phase: 'Phase',
  processGroup: 'Process Group',
};

/** Role columns, in the precedence used when one role is named under several. */
const ROLE_COLUMNS: Array<{ code: ChipCode; columns: string[] }> = [
  { code: 'A', columns: ['Accountable'] },
  { code: 'R', columns: ['Responsible'] },
  // The panel's list calls this column "Supports"; the import mapping calls it "Support".
  { code: 'S', columns: ['Supports', 'Support'] },
  { code: 'C', columns: ['Consulted', 'Consult'] },
  { code: 'I', columns: ['Informed', 'Inform'] },
];

/**
 * Role groups for the Matrix band row. The list has no such column, so the
 * grouping lives here, in display order. A role not named below lands in
 * "Other" after all of these.
 */
const ROLE_GROUPS: Array<{ group: string; roles: string[] }> = [
  { group: 'Project Leadership', roles: ['Managing Director', 'General Manager', 'Operations Manager', 'Project Delivery Manager'] },
  { group: 'Key Project Team', roles: ['Project Manager', 'Project Engineer', 'Design Lead', 'Construction Lead', 'Site Lead', 'Commissioning Lead'] },
  { group: 'Project Support', roles: ['Project Administrator', 'Health & Safety Lead', 'Environment Lead', 'Quality Lead', 'Operations Lead', 'Procurement/Contract Lead', 'Quality Inspector', 'Permit Issuer', 'Portfolio Planner'] },
  { group: 'External', roles: ['Designer(s)', 'Supplier(s)', 'Constructor(s)', 'Operator(s)'] },
  { group: 'Client', roles: ['Sponsor', "Client's Engineer"] },
  { group: 'Functional Leadership', roles: ['Engineering Manager', 'Design Team Lead', 'Construction Manager', 'Construction Team Lead', 'Delivery/Field Manager', 'Operations Team Lead', 'Quality Inspection Team Lead'] },
];
const OTHER_GROUP = 'Other';

interface IListItemsResponse {
  value: Array<Record<string, unknown>>;
  'odata.nextLink'?: string;
  '@odata.nextLink'?: string;
}

const PAGE_SIZE = 5000;

const odata = (value: string): string => value.replace(/'/g, "''");

const asString = (value: unknown): string =>
  value === null || value === undefined ? '' : String(value).trim();

const asArray = <T>(raw: unknown): T[] => {
  if (Array.isArray(raw)) { return raw as T[]; }
  if (raw && typeof raw === 'object') {
    const results = (raw as { results?: unknown }).results;
    if (Array.isArray(results)) { return results as T[]; }
  }
  return [];
};

/** A lookup/person column needs $expand and a projection, or SharePoint 400s. */
const fieldQuery = (field: IFieldInfo): { select: string[]; expand: string[] } => {
  if (!isLookupType(field.typeAsString)) {
    return { select: [field.internalName], expand: [] };
  }
  const target = field.lookupField && field.lookupField !== 'Id' ? field.lookupField : 'Title';
  return {
    select: (target === 'Id' ? ['Id'] : ['Id', target]).map((p) => `${field.internalName}/${p}`),
    expand: [field.internalName],
  };
};

const lookupTexts = (raw: unknown, field: IFieldInfo): string[] => {
  const target = field.lookupField && field.lookupField !== 'Id' ? field.lookupField : 'Title';
  const rows = asArray<Record<string, unknown>>(raw);
  const items = rows.length > 0
    ? rows
    : (raw && typeof raw === 'object' ? [raw as Record<string, unknown>] : []);
  return items
    .map((item) => asString(item[target] !== undefined ? item[target] : item.Title))
    .filter(Boolean);
};

/** Role names: straighten curly apostrophes, drop a trailing "*", tidy spaces. */
const cleanRole = (name: string): string =>
  name.replace(/[‘’]/g, "'").replace(/\s+/g, ' ').replace(/\*+$/, '').trim();

const roleKey = (name: string): string => cleanRole(name).toLowerCase();

/**
 * Semicolons and newlines only. Commas stay put: role names can contain one, and
 * shattering a single role in two is worse than leaving a comma list as one role.
 */
const splitHolders = (raw: string): string[] =>
  (raw || '').split(/[;\r\n]+/).map(cleanRole).filter(Boolean);

/** "2.1.3.10" -> [2,1,3,10]; non-numeric segments sort as 0. */
const codeParts = (code: string): number[] =>
  code.split('.').map((p) => { const n = parseInt(p, 10); return isNaN(n) ? 0 : n; });

const compareCodes = (a: string, b: string): number => {
  // A blank code has nothing to order by, so it goes last rather than first.
  if (!a && !b) { return 0; }
  if (!a) { return 1; }
  if (!b) { return -1; }
  const pa = codeParts(a);
  const pb = codeParts(b);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const d = (pa[i] === undefined ? -1 : pa[i]) - (pb[i] === undefined ? -1 : pb[i]);
    if (d !== 0) { return d; }
  }
  return 0;
};

/** A role named under several codes shows as the first of these it holds
 *  (A and R together are handled separately, as AR). */
const PRECEDENCE: ChipCode[] = ['A', 'R', 'S', 'C', 'I'];

interface IRawRow {
  id: number;
  code: string;
  parentCode: string;
  level: number;
  title: string;
  phase: string;
  group: string;
  /** role key -> codes it holds on this row */
  held: { [roleKey: string]: { [code: string]: true } };
}

export class MasterRasciService {
  private _fieldsPromise: Promise<ListFieldMap> | undefined;

  public constructor(
    private readonly _client: SPHttpClient,
    private readonly _webUrl: string,
    private readonly _listTitle: string
  ) {}

  private _fields(): Promise<ListFieldMap> {
    if (!this._fieldsPromise) {
      this._fieldsPromise = loadFieldMap(this._client, this._webUrl, this._listTitle)
        .catch((err) => {
          this._fieldsPromise = undefined; // let a Retry actually retry
          throw err;
        });
    }
    return this._fieldsPromise;
  }

  public async load(): Promise<IMasterRasciData> {
    const fields = await this._fields();

    const select: string[] = ['Id'];
    const expand: string[] = [];
    const add = (field: IFieldInfo | undefined): void => {
      if (!field) { return; }
      const parts = fieldQuery(field);
      parts.select.forEach((s) => { if (select.indexOf(s) === -1) { select.push(s); } });
      parts.expand.forEach((e) => { if (expand.indexOf(e) === -1) { expand.push(e); } });
    };

    // Title is required; a typo in the list name should say so, not render nothing.
    add(fields.requireInfo(COL.title));
    // Everything else is optional — $select on a missing column is a hard 400, so
    // each goes through info(). Phase, group and level can be derived from the code.
    [COL.processCode, COL.parentProcessCode, COL.level, COL.phase, COL.processGroup].forEach(
      (c) => add(fields.info(c)));

    const roleInfos: Array<{ code: ChipCode; field: IFieldInfo }> = [];
    ROLE_COLUMNS.forEach((rc) => {
      for (const name of rc.columns) {
        const field = fields.info(name);
        if (field) { add(field); roleInfos.push({ code: rc.code, field }); return; }
      }
    });
    if (roleInfos.length === 0) {
      throw new Error(
        `The "${this._listTitle}" list has none of the role columns ` +
        `(Accountable, Responsible, Supports, Consulted, Informed).`
      );
    }

    let url: string | undefined =
      `${this._webUrl}/_api/web/lists/getByTitle('${odata(this._listTitle)}')/items` +
      `?$select=${select.join(',')}` +
      (expand.length ? `&$expand=${expand.join(',')}` : '') +
      `&$top=${PAGE_SIZE}`;

    const items: Array<Record<string, unknown>> = [];
    while (url) {
      const body: IListItemsResponse = await this._getJson<IListItemsResponse>(url);
      items.push(...body.value);
      url = body['odata.nextLink'] || body['@odata.nextLink'];
    }

    return this._build(items, fields, roleInfos);
  }

  private _build(
    items: Array<Record<string, unknown>>,
    fields: ListFieldMap,
    roleInfos: Array<{ code: ChipCode; field: IFieldInfo }>
  ): IMasterRasciData {
    const text = (item: Record<string, unknown>, col: string): string => {
      const field = fields.info(col);
      if (!field) { return ''; }
      const raw = item[field.internalName];
      return isLookupType(field.typeAsString) ? lookupTexts(raw, field).join('; ') : asString(raw);
    };

    const holders = (item: Record<string, unknown>, field: IFieldInfo): string[] => {
      const raw = item[field.internalName];
      if (isLookupType(field.typeAsString)) { return lookupTexts(raw, field).map(cleanRole); }
      const multi = asArray<unknown>(raw); // multi-choice column
      if (multi.length > 0) { return multi.map(asString).map(cleanRole).filter(Boolean); }
      return splitHolders(asString(raw));
    };

    const displayName: { [key: string]: string } = {};
    const raws: IRawRow[] = items.map((item) => {
      const code = text(item, COL.processCode);
      const levelRaw = parseInt(text(item, COL.level), 10);
      const held: IRawRow['held'] = {};
      roleInfos.forEach((ri) => {
        holders(item, ri.field).forEach((name) => {
          const key = roleKey(name);
          if (!displayName[key]) { displayName[key] = name; }
          (held[key] = held[key] || {})[ri.code] = true;
        });
      });
      const parts = code ? code.split('.') : [];
      return {
        id: Number(item.Id),
        code,
        parentCode: text(item, COL.parentProcessCode),
        // Level column when filled in, else the depth of the code: 2.1.3 is L3.
        level: isNaN(levelRaw) ? parts.length : levelRaw,
        title: text(item, COL.title),
        phase: text(item, COL.phase) || (parts.length > 0 ? parts[0] : ''),
        group: text(item, COL.processGroup) || (parts.length > 1 ? parts.slice(0, 2).join('.') : ''),
        held,
      };
    });

    raws.sort((a, b) => compareCodes(a.code, b.code) || a.id - b.id);

    const byCode: { [code: string]: IRawRow } = {};
    const parentsWithChildren: { [code: string]: true } = {};
    raws.forEach((r) => { if (r.code) { byCode[r.code] = r; } });
    raws.forEach((r) => {
      if (r.level >= 4) {
        const parent = r.parentCode || r.code.split('.').slice(0, -1).join('.');
        if (parent) { parentsWithChildren[parent] = true; }
      }
    });

    // Roles seen in the data, in display order: known roles by group, then
    // anything unrecognised alphabetically under "Other".
    const roles = this._orderRoles(displayName);
    const roleIndex: { [key: string]: number } = {};
    roles.forEach((r, i) => { roleIndex[roleKey(r.name)] = i; });

    const rows: IRasciActivity[] = [];
    raws.forEach((r) => {
      const hasRoles = Object.keys(r.held).length > 0;
      // An L3 row that only introduces L4 children is a heading: its title is
      // already shown as the subtitle of each child, so listing it too would
      // duplicate it. One that carries roles of its own stays.
      if (r.level <= 3 && r.code && parentsWithChildren[r.code] && !hasRoles) { return; }

      const assignments: IRasciActivity['assignments'] = {};
      Object.keys(r.held).forEach((key) => {
        const codes = r.held[key];
        let code: RasciCode | undefined;
        if (codes.A && codes.R) { code = 'AR'; } else {
          for (const c of PRECEDENCE) { if (codes[c]) { code = c; break; } }
        }
        if (code) { assignments[roleIndex[key]] = code; }
      });

      let process = r.title;
      let activity = '';
      if (r.level >= 4) {
        const parentCode = r.parentCode || r.code.split('.').slice(0, -1).join('.');
        const parent = parentCode ? byCode[parentCode] : undefined;
        activity = r.title;
        process = parent ? parent.title : '';
      }

      rows.push({
        id: r.id,
        phase: r.phase,
        group: r.group,
        process,
        activity,
        assignments,
      });
    });

    return { roles, rows };
  }

  private _orderRoles(displayName: { [key: string]: string }): IRasciRole[] {
    const remaining: { [key: string]: boolean } = {};
    Object.keys(displayName).forEach((k) => { remaining[k] = true; });

    const roles: IRasciRole[] = [];
    ROLE_GROUPS.forEach((g) => {
      g.roles.forEach((known) => {
        const key = roleKey(known);
        if (remaining[key]) {
          // Show the list's own spelling, not the canonical one.
          roles.push({ name: displayName[key], group: g.group });
          delete remaining[key];
        }
      });
    });

    Object.keys(remaining)
      .map((k) => displayName[k])
      .sort((a, b) => a.localeCompare(b))
      .forEach((name) => roles.push({ name, group: OTHER_GROUP }));

    return roles;
  }

  private async _getJson<T>(url: string): Promise<T> {
    const response: SPHttpClientResponse = await this._client.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      let detail = '';
      try {
        const err = await response.json();
        detail = err && err.error && err.error.message
          ? `: ${err.error.message.value || err.error.message}`
          : '';
      } catch { /* body was not JSON — the status alone will have to do */ }
      throw new Error(`Couldn't read "${this._listTitle}" (HTTP ${response.status})${detail}`);
    }
    return response.json();
  }
}
