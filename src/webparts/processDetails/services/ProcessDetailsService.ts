/* eslint-disable @rushstack/no-new-null --
 * null means "this code has no rows in the list" — a real, renderable state
 * distinct from undefined, which here means "not fetched yet". The component
 * prop is typed IProcessStep | null to match. */
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IProcessStep, IOutputItem, ISubStep } from '../components/IProcessDetailsProps';
import { ListFieldMap, loadFieldMap, IFieldInfo, isLookupType } from './spFieldNames';

/**
 * Reads a process and its L4 variants from the "Process Details" list, with
 * artefact logos resolved from the "System Master" list.
 *
 * Everything the component needs is assembled here; ProcessDetails.tsx performs
 * no lookups of its own.
 */

export const OUTPUT_SLOTS = 8; // the source data only ever populates Output 1..8

/** Column display names. Only these need to match the lists. */
const COL = {
  title: 'Title',
  processCode: 'ProcessCode',
  parentProcessCode: 'ParentProcessCode',
  processGroup: 'Process Group',
  phase: 'Phase',
  level: 'Level',
  purpose: 'Purpose',
  inputs: 'Inputs',
  outputName: (i: number): string => `Output ${i} - Name`,
  outputSystem: (i: number): string => `Output ${i} - System`,
  outputLink: (i: number): string => `Output ${i} - Link`,
  system: 'System',
  logo: 'Logo',
};

export interface IDiagnostics {
  processFields: IFieldInfo[];
  systemFields: IFieldInfo[];
  systemLogos: Array<{ system: string; logoUrl: string }>;
  unresolvedSystems: string[];
  lastQuery?: string;
  lastRowCount?: number;
}

interface IListItemsResponse {
  value: Array<Record<string, unknown>>;
}

/** One row of a list item's AttachmentFiles collection. */
interface IAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

/**
 * Expanded collections come back as a bare array under some OData metadata
 * levels and as { results: [...] } under others. Normalise both.
 */
const asArray = <T>(raw: unknown): T[] => {
  if (Array.isArray(raw)) { return raw as T[]; }
  if (raw && typeof raw === 'object') {
    const results = (raw as { results?: unknown }).results;
    if (Array.isArray(results)) { return results as T[]; }
  }
  return [];
};

/** "2 Plan" -> "Plan", "2.1 Establish" -> "Establish". */
export const stripLeadingCode = (value: string): string =>
  (value || '').replace(/^[\d.]+\s*/, '').trim();

const odata = (value: string): string => value.replace(/'/g, "''");

const asString = (value: unknown): string =>
  value === null || value === undefined ? '' : String(value).trim();

/**
 * The Logo column can be an Image column, a Hyperlink column, or plain text,
 * and the logic doc flags its real type as unconfirmed. Handle all of them and
 * warn rather than throw on anything else — one malformed row should not take
 * the whole panel down.
 */
const extractLogoUrl = (
  raw: unknown,
  webUrl: string,
  attachments: IAttachment[]
): string | undefined => {
  if (raw === null || raw === undefined || raw === '') { return undefined; }

  const origin = webUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');
  const absolute = (path: string): string =>
    /^https?:\/\//i.test(path) ? path : origin + path;

  // Hyperlink column, already deserialised into { Url, Description }.
  if (typeof raw === 'object') {
    const url = (raw as { Url?: string }).Url;
    return url ? absolute(url) : undefined;
  }

  if (typeof raw === 'string') {
    const text = raw.trim();

    // Image column: a JSON string, in one of three shapes.
    if (text.charAt(0) === '{') {
      let parsed: { serverUrl?: string; serverRelativeUrl?: string; Url?: string; fileName?: string };
      try {
        parsed = JSON.parse(text);
      } catch {
        console.warn('[Process Details] Logo value looked like JSON but would not parse:', text);
        return undefined;
      }

      // (a) image stored in site assets — carries its own path
      if (parsed.serverRelativeUrl) {
        return (parsed.serverUrl || origin) + parsed.serverRelativeUrl;
      }
      if (parsed.Url) { return absolute(parsed.Url); }

      // (b) image stored as a list item attachment ("Reserved_ImageAttachment_…").
      // Only the file name is recorded, so the URL has to come from the item's
      // own AttachmentFiles collection.
      if (parsed.fileName) {
        const wanted = parsed.fileName.toLowerCase();
        const hit = attachments.filter((a) => (a.FileName || '').toLowerCase() === wanted)[0];
        if (hit && hit.ServerRelativeUrl) { return absolute(hit.ServerRelativeUrl); }
        console.warn(
          `[Process Details] Logo attachment "${parsed.fileName}" is referenced but was not ` +
          `found in the item's attachments.`
        );
        return undefined;
      }

      console.warn('[Process Details] Unrecognised Logo JSON shape:', text);
      return undefined;
    }

    if (/^https?:\/\//i.test(text) || text.charAt(0) === '/') { return absolute(text); }
  }

  console.warn('[Process Details] Unrecognised Logo column value:', raw);
  return undefined;
};

/**
 * A lookup or person column cannot be read with a plain $select — SharePoint
 * answers HTTP 400 and tells you to $expand it. These two helpers keep that
 * detail in one place so every column is queried correctly whatever its type.
 */
interface IQueryParts { select: string[]; expand: string[] }

const fieldQuery = (field: IFieldInfo): IQueryParts => {
  if (!isLookupType(field.typeAsString)) {
    return { select: [field.internalName], expand: [] };
  }
  // Id is always projectable and is what we match on; the display field is only
  // needed for lookups whose text we actually show (Phase, Process Group).
  const target = field.lookupField && field.lookupField !== 'Id' ? field.lookupField : 'Title';
  const props = target === 'Id' ? ['Id'] : ['Id', target];
  return {
    select: props.map((p) => `${field.internalName}/${p}`),
    expand: [field.internalName],
  };
};

/** Reads an expanded lookup back out of a row. */
const lookupValue = (raw: unknown, field: IFieldInfo): { id?: number; text: string } => {
  if (raw === null || raw === undefined) { return { text: '' }; }
  if (typeof raw !== 'object') { return { text: asString(raw) }; }

  const obj = raw as Record<string, unknown>;
  // Multi-value lookups arrive as { results: [...] }; we only ever want the first.
  const results = obj.results as Array<Record<string, unknown>> | undefined;
  const first = Array.isArray(results) ? results[0] : obj;
  if (!first) { return { text: '' }; }

  const target = field.lookupField && field.lookupField !== 'Id' ? field.lookupField : 'Title';
  const text = asString(first[target] !== undefined ? first[target] : first.Title);
  const id = Number(first.Id);
  return { id: isNaN(id) ? undefined : id, text };
};

/** System name + logo, indexed both ways so a lookup can match on id. */
interface ISystemCatalogue {
  byId: Map<number, { name: string; logoUrl: string }>;
  byName: Map<string, string>;
}

/** One row of the Process Details list, before base/variant splitting. */
interface IRecord {
  processCode: string;
  parentProcessCode: string;
  title: string;
  phase: string;
  processGroup: string;
  level: number;
  purpose: string;
  inputs: string[];
  outputs: IOutputItem[];
}

export class ProcessDetailsService {
  private _fieldsPromise: Promise<{ process: ListFieldMap; system: ListFieldMap }> | undefined;
  private _logosPromise: Promise<ISystemCatalogue> | undefined;
  private readonly _cache = new Map<string, IProcessStep | null>();
  private readonly _warnedSystems = new Set<string>();

  private _diagnostics: IDiagnostics = {
    processFields: [], systemFields: [], systemLogos: [], unresolvedSystems: [],
  };

  public constructor(
    private readonly _client: SPHttpClient,
    private readonly _webUrl: string,
    private readonly _processListTitle: string,
    private readonly _systemListTitle: string
  ) {}

  public get diagnostics(): IDiagnostics {
    return this._diagnostics;
  }

  // Both lookups are cached for the lifetime of the web part instance — the
  // schema and the system catalogue do not change between clicks.
  private _fields(): Promise<{ process: ListFieldMap; system: ListFieldMap }> {
    if (!this._fieldsPromise) {
      this._fieldsPromise = Promise.all([
        loadFieldMap(this._client, this._webUrl, this._processListTitle),
        loadFieldMap(this._client, this._webUrl, this._systemListTitle),
      ]).then(([process, system]) => {
        this._diagnostics.processFields = process.all();
        this._diagnostics.systemFields = system.all();
        return { process, system };
      }).catch((err) => {
        this._fieldsPromise = undefined; // let a Retry actually retry
        throw err;
      });
    }
    return this._fieldsPromise;
  }

  private _logoMap(): Promise<ISystemCatalogue> {
    if (!this._logosPromise) {
      this._logosPromise = this._loadLogoMap().catch((err) => {
        this._logosPromise = undefined;
        throw err;
      });
    }
    return this._logosPromise;
  }

  private async _loadLogoMap(): Promise<ISystemCatalogue> {
    const { system } = await this._fields();
    const systemCol = system.require(COL.system);
    const logoCol = system.require(COL.logo);

    // Id is needed so an "Output N - System" lookup can be matched by item id
    // rather than by display text. AttachmentFiles is needed because a modern
    // Image column often stores the picture as an item attachment and records
    // only its file name — the URL lives on the attachment.
    const base =
      `${this._webUrl}/_api/web/lists/getByTitle('${odata(this._systemListTitle)}')/items` +
      `?$top=500&$select=Id,${systemCol},${logoCol}`;
    const withAttachments =
      `${base},AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl&$expand=AttachmentFiles`;

    // Lists with attachments disabled reject the expand outright, and their
    // images will be site-asset URLs anyway — so fall back rather than fail.
    let body: IListItemsResponse;
    try {
      body = await this._getJson<IListItemsResponse>(withAttachments, this._systemListTitle);
    } catch {
      body = await this._getJson<IListItemsResponse>(base, this._systemListTitle);
    }

    const catalogue: ISystemCatalogue = { byId: new Map(), byName: new Map() };
    const listed: Array<{ system: string; logoUrl: string }> = [];
    for (const item of body.value) {
      const name = asString(item[systemCol]);
      const attachments = asArray<IAttachment>(item.AttachmentFiles);
      const logoUrl = extractLogoUrl(item[logoCol], this._webUrl, attachments);
      const id = Number(item.Id);
      if (name && logoUrl) {
        if (!isNaN(id)) { catalogue.byId.set(id, { name, logoUrl }); }
        catalogue.byName.set(name.toLowerCase(), logoUrl);
        listed.push({ system: name, logoUrl });
      } else if (name) {
        console.warn(`[Process Details] "${this._systemListTitle}" row "${name}" has no usable Logo value.`);
      }
    }
    this._diagnostics.systemLogos = listed;
    return catalogue;
  }

  /**
   * Loads one process. Returns null when the code has no rows at all, which is
   * a legitimate state (content not entered yet), not an error.
   */
  public async loadProcess(code: string): Promise<IProcessStep | null> {
    const cached = this._cache.get(code);
    if (cached !== undefined) { return cached; }

    const [{ process }, logos] = await Promise.all([this._fields(), this._logoMap()]);

    const codeCol = process.require(COL.processCode);
    const parentCol = process.require(COL.parentProcessCode);

    const select: string[] = ['Id'];
    const expand: string[] = [];

    // Every column goes through fieldQuery, because any of them could be a
    // lookup — "Output N - System" points at the systems list, and Phase or
    // Process Group may well be lookups too on some sites.
    const add = (field: IFieldInfo | undefined): void => {
      if (!field) { return; }
      const parts = fieldQuery(field);
      for (const s of parts.select) { if (select.indexOf(s) === -1) { select.push(s); } }
      for (const e of parts.expand) { if (expand.indexOf(e) === -1) { expand.push(e); } }
    };

    for (const col of [
      COL.processCode, COL.parentProcessCode, COL.title,
      COL.phase, COL.processGroup, COL.purpose, COL.inputs,
    ]) {
      add(process.requireInfo(col));
    }
    const levelCol = process.find(COL.level);
    add(process.info(COL.level));

    // Output slots are optional — a list with only four populated columns is
    // fine, and $select on a missing column is a hard 400.
    for (let i = 1; i <= OUTPUT_SLOTS; i++) {
      add(process.info(COL.outputName(i)));
      add(process.info(COL.outputSystem(i)));
      add(process.info(COL.outputLink(i)));
    }

    const url =
      `${this._webUrl}/_api/web/lists/getByTitle('${odata(this._processListTitle)}')/items` +
      `?$select=${select.join(',')}` +
      (expand.length ? `&$expand=${expand.join(',')}` : '') +
      `&$filter=${codeCol} eq '${odata(code)}' or ${parentCol} eq '${odata(code)}'` +
      `&$orderby=${codeCol}&$top=200`;

    this._diagnostics.lastQuery = url;
    const body = await this._getJson<IListItemsResponse>(url, this._processListTitle);
    this._diagnostics.lastRowCount = body.value.length;

    const records = body.value.map((item) =>
      this._mapRecord(item, process, logos, levelCol)
    );

    const step = this._assemble(code, records);
    this._cache.set(code, step);
    return step;
  }

  /**
   * Splits base from variants on code identity rather than the Level column.
   * Level is only a sanity check: a blank or mistyped Level would make
   * `find(r => r.level === 3)` silently return nothing, whereas the code
   * relationship is unambiguous by construction of the query.
   */
  private _assemble(code: string, records: IRecord[]): IProcessStep | null {
    const base = records.filter((r) => r.processCode === code)[0];
    if (!base) { return null; }

    if (base.level && base.level !== 3) {
      console.warn(`[Process Details] ${code} is the base row but its Level is ${base.level}, not 3.`);
    }

    const subSteps: ISubStep[] = records
      .filter((r) => r.parentProcessCode === code && r.processCode !== code)
      .map((r) => ({ code: r.processCode, title: r.title, outputs: r.outputs }));

    return {
      code: base.processCode,
      title: base.title,
      // Derived at render time from Phase and Process Group, never stored.
      breadcrumb: [
        stripLeadingCode(base.phase).toUpperCase(),
        stripLeadingCode(base.processGroup),
      ].filter(Boolean).join(' › '),
      purpose: base.purpose,
      inputs: base.inputs,
      // Only used when there are no sub-steps; outputs otherwise live on each L4.
      outputs: base.outputs,
      subSteps,
    };
  }

  private _mapRecord(
    item: Record<string, unknown>,
    fields: ListFieldMap,
    logos: ISystemCatalogue,
    levelCol: string | undefined
  ): IRecord {
    /** Text value of a column, transparently unwrapping an expanded lookup. */
    const get = (col: string): string => {
      const field = fields.info(col);
      if (!field) { return ''; }
      const raw = item[field.internalName];
      return isLookupType(field.typeAsString) ? lookupValue(raw, field).text : asString(raw);
    };

    const outputs: IOutputItem[] = [];
    for (let i = 1; i <= OUTPUT_SLOTS; i++) {
      const text = get(COL.outputName(i));
      if (!text) { continue; } // an empty slot, not a gap to report

      // The system column is a lookup into the systems list on this site, so
      // match on item id — exact, and immune to case or trailing spaces. Falls
      // back to name matching when the column is plain text instead.
      const sysField = fields.info(COL.outputSystem(i));
      let system = '';
      let logoUrl: string | undefined;
      if (sysField) {
        const raw = item[sysField.internalName];
        if (isLookupType(sysField.typeAsString)) {
          const v = lookupValue(raw, sysField);
          const hit = v.id !== undefined ? logos.byId.get(v.id) : undefined;
          system = hit ? hit.name : v.text;
          logoUrl = hit ? hit.logoUrl : (v.text ? logos.byName.get(v.text.toLowerCase()) : undefined);
        } else {
          system = asString(raw);
          logoUrl = system ? logos.byName.get(system.toLowerCase()) : undefined;
        }
      }

      const linkField = fields.info(COL.outputLink(i));
      const rawLink = linkField ? item[linkField.internalName] : undefined;
      const href = typeof rawLink === 'object' && rawLink !== null
        ? asString((rawLink as { Url?: string }).Url)
        : asString(rawLink);

      if (system && !logoUrl && !this._warnedSystems.has(system.toLowerCase())) {
        this._warnedSystems.add(system.toLowerCase());
        this._diagnostics.unresolvedSystems.push(system);
        console.warn(
          `[Process Details] No "${this._systemListTitle}" entry (or no logo) for system "${system}" ` +
          `— the output will render without an icon.`
        );
      }

      outputs.push({ text, system, logoUrl, href });
    }

    const inputsRaw = get(COL.inputs);

    return {
      processCode: get(COL.processCode),
      parentProcessCode: get(COL.parentProcessCode),
      title: get(COL.title),
      phase: get(COL.phase),
      processGroup: get(COL.processGroup),
      level: levelCol ? Number(item[levelCol]) || 0 : 0,
      purpose: get(COL.purpose),
      inputs: inputsRaw ? inputsRaw.split(';').map((s) => s.trim()).filter(Boolean) : [],
      outputs,
    };
  }

  private async _getJson<T>(url: string, listTitle: string): Promise<T> {
    const response: SPHttpClientResponse = await this._client.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      let detail = '';
      try {
        const err = await response.json();
        detail = err && err.error && err.error.message ? `: ${err.error.message.value || err.error.message}` : '';
      } catch { /* body was not JSON — the status alone will have to do */ }
      throw new Error(`Couldn't read "${listTitle}" (HTTP ${response.status})${detail}`);
    }
    return response.json();
  }
}
