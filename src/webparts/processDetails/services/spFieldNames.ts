import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Resolves a list's *internal* field names from its display names at runtime.
 *
 * SharePoint mangles spaces and hyphens in internal names, so a column shown as
 * "Output 1 - Name" is really `Output_x0020_1_x0020__x002d__x0020_Name` — and
 * only if it was created with that display name in the first place; a renamed
 * column keeps whatever internal name it was born with. Guessing is therefore
 * unreliable in both directions, and OData queries must use the internal name.
 *
 * Asking the Fields API once per session costs one request and removes the
 * guesswork entirely, including after a column is renamed.
 */

export interface IFieldInfo {
  internalName: string;
  displayName: string;
  typeAsString: string;
  /** Lookup/User columns only: internal name of the field shown from the target
   *  list (usually 'Title'). Needed to project a value out of an $expand. */
  lookupField?: string;
}

interface IFieldsResponse {
  value: Array<{ Title: string; InternalName: string; TypeAsString: string; LookupField?: string }>;
}

/** Column types that cannot be read with a plain $select — they need $expand
 *  and a projection like `Field/Title`, or SharePoint returns HTTP 400. */
export const isLookupType = (typeAsString: string): boolean =>
  typeAsString === 'Lookup' || typeAsString === 'LookupMulti' ||
  typeAsString === 'User' || typeAsString === 'UserMulti';

/** "Output 1 - Name" and "Output1Name" both normalise to "output1name". */
const normalise = (name: string): string => name.toLowerCase().replace(/[^a-z0-9]/g, '');

export class ListFieldMap {
  private readonly _byName: Map<string, IFieldInfo>;

  public constructor(public readonly listTitle: string, fields: IFieldInfo[]) {
    this._byName = new Map<string, IFieldInfo>();
    for (const field of fields) {
      // Display name wins on collision, but index the internal name too so a
      // column whose display name was changed after creation still resolves.
      this._byName.set(normalise(field.displayName), field);
      const internal = normalise(field.internalName);
      if (!this._byName.has(internal)) { this._byName.set(internal, field); }
    }
  }

  /** Internal name for a column, or undefined when the list has no such column. */
  public find(displayName: string): string | undefined {
    const hit = this._byName.get(normalise(displayName));
    return hit ? hit.internalName : undefined;
  }

  /** Full metadata for a column — the caller needs the type to know whether the
   *  column can be plain-$selected or has to be $expanded. */
  public info(displayName: string): IFieldInfo | undefined {
    return this._byName.get(normalise(displayName));
  }

  /** As info(), but throws a message worth showing to a page author. */
  public requireInfo(displayName: string): IFieldInfo {
    const hit = this.info(displayName);
    if (!hit) {
      throw new Error(
        `The "${this.listTitle}" list has no column called "${displayName}". ` +
        `Check the column name, or the list name in the web part properties.`
      );
    }
    return hit;
  }

  /** As find(), but throws a message worth showing to a page author. */
  public require(displayName: string): string {
    const internal = this.find(displayName);
    if (!internal) {
      throw new Error(
        `The "${this.listTitle}" list has no column called "${displayName}". ` +
        `Check the column name, or the list name in the web part properties.`
      );
    }
    return internal;
  }

  /** Every resolved column, for the diagnostics panel. */
  public all(): IFieldInfo[] {
    const seen = new Set<string>();
    const out: IFieldInfo[] = [];
    this._byName.forEach((field) => {
      if (!seen.has(field.internalName)) {
        seen.add(field.internalName);
        out.push(field);
      }
    });
    return out.sort((a, b) => a.displayName.localeCompare(b.displayName));
  }
}

const BASE_SELECT = 'Title,InternalName,TypeAsString';

export const loadFieldMap = async (
  client: SPHttpClient,
  webUrl: string,
  listTitle: string
): Promise<ListFieldMap> => {
  const base =
    `${webUrl}/_api/web/lists/getByTitle('${listTitle.replace(/'/g, "''")}')/fields` +
    `?$filter=Hidden eq false&$select=`;

  const fetch = async (select: string): Promise<SPHttpClientResponse> =>
    client.get(base + select, SPHttpClient.configurations.v1);

  // LookupField exists only on lookup/user fields. Most tenants return null for
  // the rest, but some reject the whole query — so fall back to the base select
  // rather than failing outright. Losing it only means assuming 'Title'.
  let response = await fetch(`${BASE_SELECT},LookupField`);
  if (!response.ok && response.status === 400) {
    response = await fetch(BASE_SELECT);
  }

  if (!response.ok) {
    throw new Error(
      response.status === 404
        ? `No list called "${listTitle}" on this site.`
        : `Couldn't read the columns of "${listTitle}" (HTTP ${response.status}).`
    );
  }

  const body: IFieldsResponse = await response.json();
  return new ListFieldMap(
    listTitle,
    body.value.map((f) => ({
      internalName: f.InternalName,
      displayName: f.Title,
      typeAsString: f.TypeAsString,
      lookupField: f.LookupField || undefined,
    }))
  );
};
