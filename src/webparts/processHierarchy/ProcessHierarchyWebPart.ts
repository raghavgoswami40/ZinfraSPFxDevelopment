import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import ProcessHierarchy, { IProcessHierarchyProps } from './components/ProcessHierarchy';
import { IPanelConfig } from './components/ProcessPanelContainer';
import { StoredPhase } from '../verticalFlow/components/IVerticalFlowProps';
import { DEFAULT_PHASES } from '../verticalFlow/components/verticalFlowData';
import { isValidHex } from '../processPanel/components/roleColors';

export interface IProcessHierarchyWebPartProps extends IPanelConfig {
  /** Phase, section and step structure. Falls back to the Vertical Flow default
   *  when the page has none saved. */
  phases?: StoredPhase[];
  /** Heading in the banner at the top. */
  bannerTitle?: string;
  showSparkles?: boolean;
}

const colorError = (value: string): string =>
  !value || isValidHex(value) ? '' : 'Enter a hex colour, for example #009DDC.';

export default class ProcessHierarchyWebPart extends BaseClientSideWebPart<IProcessHierarchyWebPartProps> {

  public render(): void {
    const p = this.properties;
    const element: React.ReactElement<IProcessHierarchyProps> = React.createElement(ProcessHierarchy, {
      phases: p.phases && p.phases.length > 0 ? p.phases : DEFAULT_PHASES,
      bannerTitle: p.bannerTitle === undefined ? 'Process Hierarchy' : p.bannerTitle,
      showSparkles: p.showSparkles !== false,
      isEditMode: this.displayMode === DisplayMode.Edit,
      onPhasesChange: (updated: StoredPhase[]) => {
        // Saved in the web part's own properties, so the edited structure
        // persists with the page.
        this.properties.phases = updated;
        this.render();
      },
      spHttpClient: this.context.spHttpClient,
      webUrl: this.context.pageContext.web.absoluteUrl,
      panelConfig: {
        processListTitle: p.processListTitle || 'Process Details',
        systemListTitle: p.systemListTitle || 'Systems Master',
        rasciListTitle: p.rasciListTitle || 'RASCI',
        pdTitle: p.pdTitle,
        showBreadcrumb: p.showBreadcrumb !== false,
        pdEmptyIcon: p.pdEmptyIcon,
        pdEmptyHeading: p.pdEmptyHeading,
        pdEmptyBody: p.pdEmptyBody,
        rasciTitle: p.rasciTitle,
        rasciEmptyIcon: p.rasciEmptyIcon,
        rasciEmptyIconL4: p.rasciEmptyIconL4,
        rasciEmptyHeadingL3: p.rasciEmptyHeadingL3,
        rasciEmptyBodyL3: p.rasciEmptyBodyL3,
        rasciEmptyHeadingL4: p.rasciEmptyHeadingL4,
        rasciEmptyBodyL4: p.rasciEmptyBodyL4,
        colorResponsible: p.colorResponsible,
        colorAccountable: p.colorAccountable,
        colorSupports: p.colorSupports,
        colorConsulted: p.colorConsulted,
        colorInformed: p.colorInformed,
      },
    });
    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(_currentTheme: IReadonlyTheme | undefined): void { /* no-op */ }

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
          header: { description: 'Process Hierarchy' },
          groups: [
            {
              groupName: 'Banner',
              groupFields: [
                PropertyPaneTextField('bannerTitle', { label: 'Title' }),
                PropertyPaneToggle('showSparkles', { label: 'Show starry animation' }),
              ],
            },
            {
              groupName: 'Lists',
              groupFields: [
                PropertyPaneTextField('processListTitle', { label: 'Process Details list' }),
                PropertyPaneTextField('systemListTitle', { label: 'Systems Master list' }),
                PropertyPaneTextField('rasciListTitle', { label: 'RASCI list' }),
              ],
            },
            {
              groupName: 'Process details panel',
              groupFields: [
                PropertyPaneTextField('pdTitle', { label: 'Header title' }),
                PropertyPaneToggle('showBreadcrumb', { label: 'Show breadcrumb' }),
                PropertyPaneTextField('rasciTitle', { label: 'RASCI header title' }),
              ],
            },
            {
              groupName: 'RASCI colours',
              groupFields: [
                PropertyPaneTextField('colorResponsible', { label: 'Responsible', onGetErrorMessage: colorError }),
                PropertyPaneTextField('colorAccountable', { label: 'Accountable', onGetErrorMessage: colorError }),
                PropertyPaneTextField('colorSupports', { label: 'Supports', onGetErrorMessage: colorError }),
                PropertyPaneTextField('colorConsulted', { label: 'Consulted', onGetErrorMessage: colorError }),
                PropertyPaneTextField('colorInformed', { label: 'Informed', onGetErrorMessage: colorError }),
              ],
            },
          ],
        },
      ],
    };
  }
}
