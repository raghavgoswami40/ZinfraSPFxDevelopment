import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import VerticalFlow from './components/VerticalFlow';
import { IVerticalFlowProps, StoredPhase } from './components/IVerticalFlowProps';
import { DEFAULT_PHASES } from './components/verticalFlowData';

export interface IVerticalFlowWebPartProps {
  urls: Record<string, string>;
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
        urls:           this.properties.urls || {},
        phases:         this.phases,
        isEditMode:     this.displayMode === DisplayMode.Edit,
        onPhasesChange: (updated: StoredPhase[]) => {
          this.properties.phases = updated;
          this.context.propertyPane.refresh();
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
    const fields: import('@microsoft/sp-property-pane').IPropertyPaneField<unknown>[] = [];

    this.phases.forEach(phase => {
      fields.push(PropertyPaneLabel(`lbl_${phase.id}`, { text: `▸ ${phase.title}` }));
      phase.sections.forEach((section, si) => {
        fields.push(PropertyPaneLabel(`lbl_${section.id}`, { text: `  ${section.title.replace(/\n/g, ' ')}` }));
        fields.push(
          PropertyPaneTextField(`urls[${section.id}_goToFlow]`, {
            label: `    Go to flow URL`,
            placeholder: 'https://',
          })
        );
        section.steps.forEach((step, idx) => {
          fields.push(
            PropertyPaneTextField(`urls[${step.id}]`, {
              label: `    Step ${idx + 1} — ${step.label}`,
              placeholder: 'https://',
            })
          );
        });
        if (si < phase.sections.length - 1) {
          fields.push(PropertyPaneLabel(`lbl_sep_${section.id}`, { text: '' }));
        }
      });
    });

    return {
      pages: [
        {
          header: { description: 'Set the SharePoint page URL for each step. To add, remove, or rename phases/sections/steps, edit the webpart directly on the page.' },
          groups: [{ groupName: 'Step URLs', groupFields: fields }],
        },
      ],
    };
  }
}
