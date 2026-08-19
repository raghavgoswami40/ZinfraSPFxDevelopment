import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import VerticalFlow from './components/VerticalFlow';
import { IVerticalFlowProps, StoredPhase } from './components/IVerticalFlowProps';
import { DEFAULT_PHASES } from './components/verticalFlowData';

export interface IVerticalFlowWebPartProps {
  phases: StoredPhase[];
}

export default class VerticalFlowWebPart extends BaseClientSideWebPart<IVerticalFlowWebPartProps> {

  private get phases(): StoredPhase[] {
    if (!this.properties.phases || this.properties.phases.length === 0) {
      this.properties.phases = JSON.parse(JSON.stringify(DEFAULT_PHASES));
    }
    return this.properties.phases;
  }

  public render(): void {
    const element: React.ReactElement<IVerticalFlowProps> = React.createElement(
      VerticalFlow,
      {
        phases:         this.phases,
        isEditMode:     this.displayMode === DisplayMode.Edit,
        onPhasesChange: (updated: StoredPhase[]) => {
          this.properties.phases = updated;
          this.render();
        },
      }
    );
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
    // No configurable properties — step/section/phase content is edited directly
    // on the page (edit mode), and steps have no links to configure here.
    return {
      pages: [
        {
          header: { description: 'Edit phases, sections, and steps directly on the page in edit mode.' },
          groups: [],
        },
      ],
    };
  }
}
