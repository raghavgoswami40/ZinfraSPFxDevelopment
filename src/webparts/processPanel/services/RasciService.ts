/* eslint-disable @rushstack/no-new-null --
 * null means "this code has no row in the RASCI list" — a real, renderable
 * state distinct from undefined, which here means "not fetched yet". */
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  IRasciRecord, RoleKey, ROLE_ORDER, ROLE_META,
} from '../components/IProcessPanelProps';
import {
  ListFieldMap, loadFieldMap, IFieldInfo, isLookupType,
} from '../../../shared/spFieldNames';

/**
 * Reads one process's RASCI from the "RASCI" list.
 *
 * The list holds L3 and L4 rows side by side, exactly like the Process Details
 * list, and is keyed the same way — ProcessCode identifies the row,
 * ParentProcessCode ties an L4 to its L3. This service only ever fetches the
 * one row it was asked for; deciding *which* code to ask for is the web part's
 * job, because that depends on the panel's L3/L4 state rather than on the data.
 */

/** Column display names. Only these need to match the list. */
const COL = {
  title: 'Title',
  processCode: 'ProcessCode',
  parentProcessCode: 'ParentProcessCode',
  level: 'Level',
};

export interface IRasciDiagnostics {
  fields: IFieldInfo[];
  lastQuery?: string;
  lastRowCount?: number;
}

interface IListItemsResponse {
  value: Array<Record<string, unknown>>;
}

const odata = (value: string): string => value.replace(/'/g, "''");

const asString = (value: unknown): string =>
  value === null || value === undefined ? '' : String(value).trim();

/**
 * Expanded collections and multi-choice columns come back as a bare array under
 * some OData metadata levels and as { results: [...] } under others.
 */
const asArray = <T>(raw: unknown): T[] => {
  if (Array.isArray(raw)) { return raw as T[]; }
  if (raw && typeof raw === 'object') {
    const results = (raw as { results?: unknown }).results;
    if (Array.isArray(results)) { return results as T[]; }
  }
  return [];
};

/**
 * Splits one role cell into individual holders.
 *
 * Semicolons and newlines only. Commas are deliberately left alone: role names
 * in this data routinely contain one ("Manager, Projects"), and splitting on it
 * would shatter a single role into two meaningless tags — a far worse failure
 * than leaving a comma-separated cell as one tag.
 */
const splitHolders = (raw: string): string[] =>
  (raw || '')
    .split(/[;\r\n]+/)
    .map((part) => part.trim())
    .filter(Boolean);

/** Reads an expanded lookup/person value back out of a row. */
const lookupTexts = (raw: unknown, field: IFieldInfo): string[] => {
  const target = field.lookupField && field.lookupField !== 'Id' ? field.lookupField : 'Title';
  const rows = asArray<Record<string, unknown>>(raw);

  // Single-valued lookups arrive as a bare object rather than a collection.
  const items = rows.length > 0
    ? rows
    : (raw && typeof raw === 'object' ? [raw as Record<string, unknown>] : []);

  return items
    .map((item) => asString(item[target] !== undefined ? item[target] : item.Title))
    .filter(Boolean);
};

/**
 * A lookup or person column cannot be read with a plain $select — SharePoint
 * answers HTTP 400 and tells you to $expand it instead.
 */
interface IQueryParts { select: string[]; expand: string[] }

const fieldQuery = (field: IFieldInfo): IQueryParts => {
  if (!isLookupType(field.typeAsString)) {
    return { select: [field.internalName], expand: [] };
  }
  const target = field.lookupField && field.lookupField !== 'Id' ? field.lookupField : 'Title';
  const props = target === 'Id' ? ['Id'] : ['Id', target];
  return {
    select: props.map((p) => `${field.internalName}/${p}`),
    expand: [field.internalName],
  };
};

export class RasciService {
  private _fieldsPromise: Promise<ListFieldMap> | undefined;
  private readonly _cache = new Map<string, IRasciRecord | null>();

  private _diagnostics: IRasciDiagnostics = { fields: [] };

  public constructor(
    private readonly _client: SPHttpClient,
    private readonly _webUrl: string,
    private readonly _listTitle: string
  ) {}

  public get diagnostics(): IRasciDiagnostics {
    return this._diagnostics;
  }

  /** Cached for the lifetime of the web part instance — a list's schema does
   *  not change between clicks. */
  private _fields(): Promise<ListFieldMap> {
    if (!this._fieldsPromise) {
      this._fieldsPromise = loadFieldMap(this._client, this._webUrl, this._listTitle)
        .then((map) => {
          this._diagnostics.fields = map.all();
          return map;
        })
        .catch((err) => {
          this._fieldsPromise = undefined; // let a Retry actually retry
          throw err;
        });
    }
    return this._fieldsPromise;
  }

  /**
   * Loads one process's RASCI. Returns null when the code has no row at all,
   * which is a legitimate state (not authored yet), not an error.
   */
  public async loadRasci(code: string): Promise<IRasciRecord | null> {
    const cached = this._cache.get(code);
    if (cached !== undefined) { return cached; }

    const fields = await this._fields();
    const codeCol = fields.require(COL.processCode);

    const select: string[] = ['Id'];
    const expand: string[] = [];

    // Every column goes through fieldQuery, because any of them could be a
    // lookup or a person column on a given site — the five role columns most
    // of all, since naming a person there is the obvious way to build this list.
    const add = (field: IFieldInfo | undefined): void => {
      if (!field) { return; }
      const parts = fieldQuery(field);
      for (const s of parts.select) { if (select.indexOf(s) === -1) { select.push(s); } }
      for (const e of parts.expand) { if (expand.indexOf(e) === -1) { expand.push(e); } }
    };

    // Required: without ProcessCode and Title there is nothing to render, and a
    // typo in the list name should say so rather than return an empty panel.
    add(fields.requireInfo(COL.processCode));
    add(fields.requireInfo(COL.title));

    // Optional: a list without ParentProcessCode or Level still renders. Level
    // is kept for a future sanity check even though nothing displays it today.
    // $select on a missing column is a hard 400, so these go through info()
    // rather than requireInfo().
    for (const col of [COL.parentProcessCode, COL.level]) {
      add(fields.info(col));
    }
    for (const role of ROLE_ORDER) {
      add(fields.info(ROLE_META[role].column));
    }

    const url =
      `${this._webUrl}/_api/web/lists/getByTitle('${odata(this._listTitle)}')/items` +
      `?$select=${select.join(',')}` +
      (expand.length ? `&$expand=${expand.join(',')}` : '') +
      `&$filter=${codeCol} eq '${odata(code)}'` +
      `&$top=2`;

    this._diagnostics.lastQuery = url;
    const body = await this._getJson<IListItemsResponse>(url);
    this._diagnostics.lastRowCount = body.value.length;

    if (body.value.length === 0) {
      this._cache.set(code, null);
      return null;
    }
    if (body.value.length > 1) {
      console.warn(
        `[RASCI] "${this._listTitle}" has more than one row with ProcessCode "${code}" ` +
        `— using the first. The list should hold one row per process.`
      );
    }

    const record = this._mapRecord(body.value[0], fields);
    this._cache.set(code, record);
    return record;
  }

  private _mapRecord(item: Record<string, unknown>, fields: ListFieldMap): IRasciRecord {
    /** Text value of a column, transparently unwrapping an expanded lookup. */
    const get = (col: string): string => {
      const field = fields.info(col);
      if (!field) { return ''; }
      const raw = item[field.internalName];
      return isLookupType(field.typeAsString)
        ? lookupTexts(raw, field).join('; ')
        : asString(raw);
    };

    /**
     * Holders for one role. Three shapes are possible and all three occur:
     *  - a person or lookup column, single or multi  -> one holder per entry
     *  - a multi-choice column                       -> one holder per choice
     *  - plain text                                  -> split on ; and newline
     */
    const holders = (col: string): string[] => {
      const field = fields.info(col);
      if (!field) { return []; }
      const raw = item[field.internalName];

      if (isLookupType(field.typeAsString)) {
        return lookupTexts(raw, field);
      }

      const multi = asArray<unknown>(raw);
      if (multi.length > 0) {
        return multi.map(asString).filter(Boolean);
      }

      return splitHolders(asString(raw));
    };

    const roles = {} as Record<RoleKey, string[]>;
    for (const role of ROLE_ORDER) {
      roles[role] = holders(ROLE_META[role].column);
    }

    const levelField = fields.info(COL.level);
    const levelRaw = levelField ? Number(item[levelField.internalName]) : NaN;

    return {
      code: get(COL.processCode),
      title: get(COL.title),
      level: isNaN(levelRaw) ? 0 : levelRaw,
      roles,
    };
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
