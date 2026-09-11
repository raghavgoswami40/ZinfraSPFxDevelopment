import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

import VerticalFlow from './components/VerticalFlow';
import { IVerticalFlowProps, StoredPhase } from './components/IVerticalFlowProps';
import { DEFAULT_PHASES } from './components/verticalFlowData';
import { resolveSelectedStep } from './stepSelection';
import { ISelectedProcessStep, SELECTED_STEP_PROPERTY_ID } from '../../shared/dynamicData';

export interface IVerticalFlowWebPartProps {
  phases: StoredPhase[];
}

export default class VerticalFlowWebPart
  extends BaseClientSideWebPart<IVerticalFlowWebPartProps>
  implements IDynamicDataCallables {

  /**
   * Which step the reader has selected. Deliberately a class field rather than
   * React state or a property-bag entry:
   *  - getPropertyValue() is called on this instance, synchronously, and has no
   *    access to hook state;
   *  - selection is transient view state, and the property bag is not writable
   *    in display mode anyway.
   * The consequence is that selection resets on page reload, which is correct.
   */
  private _selectedStepId: string | undefined = undefined;

  /** Derived payload, cached only so getPropertyValue() is cheap. Never the
   *  source of truth — _publishSelection re-derives it from phases. */
  private _selectedStep: ISelectedProcessStep | undefined = undefined;

  private get phases(): StoredPhase[] {
    if (!this.properties.phases || this.properties.phases.length === 0) {
      this.properties.phases = JSON.parse(JSON.stringify(DEFAULT_PHASES));
    }
    return this.properties.phases;
  }

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);

    // A hierarchy normally holds every phase, so its name is just "Process
    // Hierarchy". Only when one has been narrowed to a single phase is the
    // phase name added — that is the case where a page carries several of these
    // and a consumer's source dropdown would otherwise show identical entries.
    const phases = this.phases;
    const onlyPhase = phases.length === 1 ? (phases[0].title || '').replace(/\s+/g, ' ').trim() : '';
    this.context.dynamicDataSourceManager.updateMetadata({
      title: onlyPhase ? `Process Hierarchy — ${onlyPhase}` : 'Process Hierarchy',
      description: 'Publishes the process step the reader has selected in the flow diagram.',
    });

    return super.onInit();
  }

  // ── Dynamic Data source ──────────────────────────────────────────────────

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: SELECTED_STEP_PROPERTY_ID,
        title: 'Selected process step',
        description: 'The step the reader last selected, with its generated code, label, section and phase.',
      },
    ];
  }

  // Narrower than the interface's `any` on purpose: `any` is assignable both
  // ways, so this still satisfies IDynamicDataCallables. Returns undefined,
  // never null — DynamicProperty.tryGetValue() is typed `TValue | undefined`.
  public getPropertyValue(propertyId: string): ISelectedProcessStep | undefined {
    if (propertyId === SELECTED_STEP_PROPERTY_ID) {
      return this._selectedStep;
    }
    throw new Error(`VerticalFlowWebPart: unknown dynamic data property '${propertyId}'`);
  }

  /**
   * The single choke point for selection changes. Re-derives the payload from
   * the current phases, repaints, then notifies consumers — but only when the
   * payload actually changed, so a burst of edit-mode keystroke commits does
   * not broadcast a notify per keystroke.
   */
  private _publishSelection(): void {
    const next = resolveSelectedStep(this.phases, this._selectedStepId);
    const changed = JSON.stringify(next) !== JSON.stringify(this._selectedStep);

    this._selectedStep = next;
    // resolveSelectedStep returns undefined when the step no longer exists —
    // i.e. it was deleted in edit mode while selected. Drop the dangling id.
    if (!next) { this._selectedStepId = undefined; }

    // Repaint this web part first, so its own highlight updates even if a
    // consumer throws inside the synchronous notify below.
    this.render();

    if (changed) {
      this.context.dynamicDataSourceManager.notifyPropertyChanged(SELECTED_STEP_PROPERTY_ID);
    }
  }

  public render(): void {
    const element: React.ReactElement<IVerticalFlowProps> = React.createElement(
      VerticalFlow,
      {
        phases:         this.phases,
        isEditMode:     this.displayMode === DisplayMode.Edit,
        selectedStepId: this._selectedStepId,
        onStepSelect:   (stepId: string) => {
          // Clicking the selected step again clears the selection.
          this._selectedStepId = stepId === this._selectedStepId ? undefined : stepId;
          this._publishSelection();
        },
        onPhasesChange: (updated: StoredPhase[]) => {
          this.properties.phases = updated;
          // Goes through _publishSelection rather than a bare render(): editing
          // the selected step's label would otherwise leave a stale label in the
          // published payload, and deleting it a dangling stepId.
          this._publishSelection();
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
