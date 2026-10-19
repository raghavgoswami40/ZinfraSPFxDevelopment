import { SPHttpClient } from '@microsoft/sp-http';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

/**
 * The visible lists of a site (document libraries left out), as options for a property-pane
 * dropdown. `current` values are kept in the result even when the site does not have them,
 * so a saved name never silently vanishes from the dropdown.
 */
export const loadListOptions = (
  client: SPHttpClient,
  webUrl: string,
  current: string[]
): Promise<IPropertyPaneDropdownOption[]> =>
  client
    .get(`${webUrl}/_api/web/lists?$select=Title&$filter=Hidden eq false and BaseType eq 0&$orderby=Title`, SPHttpClient.configurations.v1)
    .then((r) => r.json())
    .then((j: { value?: Array<{ Title: string }> }) => {
      const titles = (j.value || []).map((l) => l.Title);
      current.forEach((c) => { if (c && titles.indexOf(c) < 0) titles.push(c); });
      return titles.map((t) => ({ key: t, text: t }));
    })
    .catch(() => current.filter(Boolean).map((t) => ({ key: t, text: t })));
