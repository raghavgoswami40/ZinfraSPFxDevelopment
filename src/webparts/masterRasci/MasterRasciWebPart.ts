import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import MasterRasci from './components/MasterRasci';
import {
  IMasterRasciProps, IMasterRasciData, IMasterRasciLabels, RasciView, DEFAULT_LABELS,
} from './components/IMasterRasciProps';
import { MasterRasciService } from './services/MasterRasciService';

export interface IMasterRasciWebPartProps {
  listTitle: string;
  documentRef: string;
  defaultView: RasciView;
  collapseAfter: number;
  labelPhase?: string;
  labelSearch?: string;
  labelGroup?: string;
  labelRole?: string;
  labelViewList?: string;
  labelViewMatrix?: string;
}

const DEFAULT_LIST = 'RASCI';
const DEFAULT_COLLAPSE_AFTER = 3;

export default class MasterRasciWebPart extends BaseClientSideWebPart<IMasterRasciWebPartProps> {

  private _service: MasterRasciService | undefined;
  private _serviceKey = '';
  private _load: (() => Promise<IMasterRasciData>) | undefined;

  private get _listTitle(): string {
    return (this.properties.listTitle || '').trim() || DEFAULT_LIST;
  }

  /**
   * The loader is rebuilt only when the list name changes. The component reloads
   * whenever it receives a new `load`, so handing it a fresh closure on every
   * render() would refetch the whole list on each property-pane keystroke.
   */
  private _getLoader(): () => Promise<IMasterRasciData> {
    const key = this._listTitle;
    if (!this._service || !this._load || this._serviceKey !== key) {
      const service = new MasterRasciService(
        this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl,
        key
      );
      this._service = service;
      this._load = () => service.load();
      this._serviceKey = key;
    }
    return this._load;
  }

  /** A blank field falls back to the default, so a label can never end up empty. */
  private get _labels(): IMasterRasciLabels {
    const p = this.properties;
    const pick = (value: string | undefined, fallback: string): string => (value || '').trim() || fallback;
    return {
      phase: pick(p.labelPhase, DEFAULT_LABELS.phase),
      search: pick(p.labelSearch, DEFAULT_LABELS.search),
      group: pick(p.labelGroup, DEFAULT_LABELS.group),
      role: pick(p.labelRole, DEFAULT_LABELS.role),
      viewList: pick(p.labelViewList, DEFAULT_LABELS.viewList),
      viewMatrix: pick(p.labelViewMatrix, DEFAULT_LABELS.viewMatrix),
    };
  }

  public render(): void {
    const collapse = Number(this.properties.collapseAfter);
    const element: React.ReactElement<IMasterRasciProps> = React.createElement(MasterRasci, {
      load: this._getLoader(),
      listTitle: this._listTitle,
      documentRef: (this.properties.documentRef || '').trim(),
      defaultView: this.properties.defaultView === 'matrix' ? 'matrix' : 'list',
      collapseAfter: collapse >= 1 ? Math.floor(collapse) : DEFAULT_COLLAPSE_AFTER,
      labels: this._labels,
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Filters and both views are used on the page itself.' },
          groups: [
            {
              groupName: 'Data',
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: 'RASCI list',
                  description: 'Title of the SharePoint list on this site that holds the RASCI.',
                  placeholder: DEFAULT_LIST,
                }),
                PropertyPaneTextField('documentRef', {
                  label: 'Document reference',
                  description: 'Shown in the footer, e.g. GS-PM-ST-001. Leave empty to hide.',
                }),
              ],
            },
            {
              groupName: 'Labels',
              groupFields: [
                PropertyPaneTextField('labelPhase', { label: 'Phase filter', placeholder: DEFAULT_LABELS.phase }),
                PropertyPaneTextField('labelSearch', { label: 'Search filter', placeholder: DEFAULT_LABELS.search }),
                PropertyPaneTextField('labelGroup', { label: 'Process group filter', placeholder: DEFAULT_LABELS.group }),
                PropertyPaneTextField('labelRole', { label: 'Role filter', placeholder: DEFAULT_LABELS.role }),
                PropertyPaneTextField('labelViewList', { label: 'First view toggle', placeholder: DEFAULT_LABELS.viewList }),
                PropertyPaneTextField('labelViewMatrix', { label: 'Second view toggle', placeholder: DEFAULT_LABELS.viewMatrix }),
              ],
            },
            {
              groupName: 'Display',
              groupFields: [
                PropertyPaneDropdown('defaultView', {
                  label: 'Default view',
                  options: [
                    { key: 'list', text: this._labels.viewList },
                    { key: 'matrix', text: this._labels.viewMatrix },
                  ],
                  selectedKey: this.properties.defaultView === 'matrix' ? 'matrix' : 'list',
                }),
                PropertyPaneSlider('collapseAfter', {
                  label: 'Roles shown per cell before "+N more"',
                  min: 1,
                  max: 12,
                  step: 1,
                  value: this.properties.collapseAfter || DEFAULT_COLLAPSE_AFTER,
                  showValue: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
