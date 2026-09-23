import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Guid, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
  DynamicDataSharedDepth,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, DynamicProperty } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { StickyPanelController, findStackFollower } from '../../shared/stickyPanel';

import ProcessDetails from './components/ProcessDetails';
import {
  IProcessDetailsProps, IProcessStep,
} from './components/IProcessDetailsProps';
import { SAMPLE_STEP_WITH_SUBSTEPS } from './components/processDetailsData';
import { ProcessDetailsService } from './services/ProcessDetailsService';
import {
  ISelectedProcessStep, SELECTED_STEP_PROPERTY_ID, VERTICAL_FLOW_COMPONENT_ID,
  IProcessFocus, PROCESS_FOCUS_PROPERTY_ID,
} from '../../shared/dynamicData';

export interface IProcessDetailsWebPartProps {
  title: string;
  headerIcon: string;
  showBreadcrumb: boolean;

  // ── Unselected state (shown before anything is clicked in the hierarchy) ──
  emptyIcon: string;
  emptyHeading: string;
  emptyBody: string;

  // The two below are deliberately not exposed in the property pane — they are
  // fixed behaviour, not authoring choices. They still take effect, from the
  // manifest defaults; re-add a PropertyPaneToggle for either to make it
  // configurable again.
  /** Accordion behaviour for L4 rows. Fixed on. */
  singleOpenSubStep: boolean;
  /** Renders the worked example instead of live data. Fixed off. */
  showSampleData: boolean;

  /** Console + in-panel report of field resolution and panel pinning. Off by
   *  default, and exposed in the property pane under Troubleshooting: when the
   *  panel pins somewhere unexpected, the console readout naming the ancestor
   *  it anchored to is the only way to tell why from outside the page. */
  showDiagnostics: boolean;
  processListTitle: string;
  systemListTitle: string;
  /** Keeps the panel on screen while the tall hierarchy beside it scrolls. */
  stickyPanel: boolean;
  /** Gap left above the pinned panel, to clear SharePoint's own sticky chrome. */
  stickyTopOffset: number;
  /** Named `selection` rather than `selectedStep` so it cannot be confused with
   *  the component prop of that name. */
  selection: DynamicProperty<ISelectedProcessStep>;
}

const DEFAULT_PROCESS_LIST = 'Process Details';
const DEFAULT_SYSTEM_LIST = 'Systems Master';

const DEFAULT_STICKY_TOP = 90;

export default class ProcessDetailsWebPart
  extends BaseClientSideWebPart<IProcessDetailsWebPartProps>
  implements IDynamicDataCallables {

  private _service: ProcessDetailsService | undefined;
  private _serviceKey = '';

  // Render state for the current code.
  private _code: string | undefined = undefined;
  private _step: IProcessStep | null = null;
  private _loading = false;
  private _error: string | undefined = undefined;
  private _notFound: string | undefined = undefined;

  /** Last diagnostics state logged, so repaints do not spam the console. */
  private _lastDiagKey = '';

  // Guards against an earlier, slower request overwriting a newer selection.
  private _requestSeq = 0;

  /** Which L4 row is expanded. Lifted out of the component so it can be
   *  published: RASCI needs the open L4, not just the selected L3. null means
   *  every row closed, which a consumer has to tell apart from "no L3
   *  selected". */
  private _openSubStepCode: string | null = null;

  /** Last published payload — what getPropertyValue() returns, and what
   *  _publishFocus() compares against so a plain repaint does not notify. */
  private _focus: IProcessFocus = { hasSubSteps: false, pending: false };

  /** Set while a notify is already queued, so several state changes in one
   *  turn collapse into a single notification. */
  private _focusNotifyQueued = false;

  /** Class-property arrow, not a method: unregisterAvailableSourcesChanged
   *  matches on function reference, so .bind() at each site would never
   *  unregister. */
  private _onAvailableSourcesChanged = (): void => { this.render(); };

  /** Pins the panel to the viewport while the page scrolls. Nothing stacks
   *  above this panel, so it has no anchor of its own; it does however give up
   *  its pin for a RASCI panel stacked beneath it when the two cannot both fit
   *  on screen — see stackFollower below. Built in onInit(), not as a field
   *  initializer — this.domElement is not guaranteed to exist yet at
   *  construction time. */
  private _sticky: StickyPanelController | undefined = undefined;

  // ── List picker for the property pane ────────────────────────────────────
  // Populated once per web part instance (not once per pane-open — a second
  // open reuses the same result) so choosing between the Process Details and
  // Systems Master lists is a dropdown of what actually exists on this site,
  // rather than a text field an author has to get exactly right by hand.
  private _listOptions: IPropertyPaneDropdownOption[] | undefined = undefined;
  private _listsPromise: Promise<IPropertyPaneDropdownOption[]> | undefined = undefined;
  // Set once a fetch has failed, so the pane falls back to plain text fields
  // rather than leaving the author stuck with a dropdown that can never
  // offer anything but whatever value is already saved.
  private _listsFetchFailed = false;

  private _loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    if (this._listsPromise) { return this._listsPromise; }

    // BaseTemplate 100 is a generic custom list — the kind Process Details and
    // Systems Master both are. Filtering to that (and out hidden lists) keeps
    // document libraries, calendars and SharePoint's own system lists out of
    // an author's way.
    const url =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$select=Title&$filter=Hidden eq false and BaseTemplate eq 100&$orderby=Title&$top=500`;

    this._listsPromise = this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) { throw new Error(`HTTP ${response.status}`); }
        return response.json();
      })
      .then((body: { value: Array<{ Title: string }> }) =>
        body.value.map((l) => ({ key: l.Title, text: l.Title }))
      )
      .catch((err) => {
        console.warn(
          '[Process Details] Could not read this site\'s lists for the property pane picker ' +
          '— falling back to plain text fields:',
          err
        );
        this._listsFetchFailed = true;
        this._listsPromise = undefined; // let a later pane-open retry
        return [];
      });

    return this._listsPromise;
  }

  /**
   * Options for one of the two list dropdowns. Always includes the currently
   * configured value even if the live fetch hasn't returned it — either
   * because it is still loading, or because the saved list was renamed,
   * deleted, or isn't a plain generic list (so the fetch legitimately
   * excludes it) — so the author's existing choice is never silently dropped
   * from the control.
   */
  private _listDropdownOptions(current: string, fallback: string): IPropertyPaneDropdownOption[] {
    const value = current || fallback;
    const loaded = this._listOptions;

    if (!loaded) {
      return [{ key: value, text: `${value} (loading site lists…)` }];
    }
    if (loaded.some((o) => o.key === value)) { return loaded; }
    return [{ key: value, text: `${value} (not found on this site)` }, ...loaded];
  }

  // Fired every time the property pane is opened. Kicks off the list fetch
  // once per instance; context.propertyPane.refresh() repaints the pane with
  // real options once it resolves, rather than a placeholder.
  protected onPropertyPaneConfigurationStart(): void {
    if (this._listOptions || this._listsFetchFailed) { return; }
    // _loadLists() already swallows its own rejections (see its .catch), so
    // this .catch() is just to satisfy no-floating-promises — it can't fire.
    this._loadLists()
      .then((options) => {
        if (options.length > 0) { this._listOptions = options; }
        this.context.propertyPane.refresh();
      })
      .catch(() => { /* unreachable — _loadLists() never rejects */ });
  }

  protected onInit(): Promise<void> {
    // Deserialization runs before onInit, so if a connection was saved the
    // property already exists and this is a no-op. It only covers an instance
    // that has never had one, where the key is absent from the property bag and
    // tryGetValue() would throw.
    if (!this.properties.selection) {
      this.properties.selection = new DynamicProperty<ISelectedProcessStep>(
        this.context.dynamicDataProvider,
        (): void => { this.render(); }
      );
    }

    // Registered as a *source* as well as a consumer: RASCI reads the focus
    // published here. Done before the first render so a RASCI web part that
    // initialises after this one finds the source already in the list.
    this.context.dynamicDataSourceManager.initializeSource(this);
    this.context.dynamicDataSourceManager.updateMetadata({
      title: this.properties.title || 'Process Details',
      description:
        'Publishes the process the reader is focused on — the selected L3, or the L4 opened within it.',
    });

    // DynamicProperty copes with a source appearing late on its own, but not
    // with one disappearing — this is what repaints the empty state when the
    // hierarchy web part is removed, instead of leaving a stale step on screen.
    this.context.dynamicDataProvider.registerAvailableSourcesChanged(this._onAvailableSourcesChanged);

    this._sticky = new StickyPanelController({
      domElement: this.domElement,
      stickyEnabled: () => this.properties.stickyPanel !== false,
      isEditMode: () => this.displayMode === DisplayMode.Edit,
      topOffset: () => (
        typeof this.properties.stickyTopOffset === 'number'
          ? this.properties.stickyTopOffset
          : DEFAULT_STICKY_TOP
      ),
      // Yields its pin to a RASCI panel stacked beneath, but only when the two
      // cannot both fit the visible band — holding it then would strand RASCI's
      // lower half below the fold where page scrolling cannot reach it. This
      // decides *whether* this panel pins, never where, so it can never lift it
      // above the top of its own section. Undefined when no RASCI panel follows
      // this one in the same column, which leaves behaviour exactly as it is
      // for a Process Details panel on its own.
      stackFollower: () => findStackFollower('rasci', this.domElement),
      diagnosticsEnabled: () => !!this.properties.showDiagnostics,
      logLabel: 'Process Details',
    });
    this._sticky.attachListeners();

    return super.onInit();
  }

  // Without this the framework will not construct the DynamicProperty for us,
  // and we would own its rehydration and disposal. 'object', not 'string' —
  // the payload is ISelectedProcessStep.
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'selection': { dynamicPropertyType: 'object' },
    };
  }

  // ── Dynamic Data source ──────────────────────────────────────────────────

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: PROCESS_FOCUS_PROPERTY_ID,
        title: 'Focused process',
        description:
          'The L3 selected in the hierarchy, or the L4 opened inside it when the L3 has sub-processes.',
      },
    ];
  }

  // Narrower than the interface's `any` on purpose — `any` is assignable both
  // ways, so this still satisfies IDynamicDataCallables.
  public getPropertyValue(propertyId: string): IProcessFocus | undefined {
    if (propertyId === PROCESS_FOCUS_PROPERTY_ID) {
      return this._focus;
    }
    throw new Error(`ProcessDetailsWebPart: unknown dynamic data property '${propertyId}'`);
  }

  /** The payload for the current render state. Pure — no side effects. */
  private _computeFocus(): IProcessFocus {
    const step = this._step;
    const hasSubSteps = !!step && step.subSteps.length > 0;

    // While the L3's row is still in flight hasSubSteps is not yet known, so
    // consumers are told to hold rather than briefly render the L3's own RASCI
    // for a step that will turn out to have L4 children.
    const pending = this._loading;

    const focus: IProcessFocus = {
      l3Code: this._code,
      l3Title: step ? step.title : undefined,
      hasSubSteps,
      pending,
    };

    if (hasSubSteps) {
      const open = this._openSubStepCode
        ? step!.subSteps.filter((sub) => sub.code === this._openSubStepCode)[0]
        : undefined;
      if (open) {
        focus.l4Code = open.code;
        focus.l4Title = open.title;
        focus.focusCode = open.code;
        focus.focusTitle = open.title;
      }
    } else if (step) {
      // A childless L3 answers for itself.
      focus.focusCode = step.code;
      focus.focusTitle = step.title;
    }

    return focus;
  }

  /**
   * Recomputes the payload and notifies consumers when it actually changed.
   *
   * The notify is deferred to a microtask because the main caller chain runs
   * *inside* this web part's render(): notifying synchronously would re-enter a
   * consumer's render from within ours. A microtask runs once the current
   * render has fully unwound, which is both safe and imperceptible.
   */
  private _publishFocus(): void {
    const next = this._computeFocus();
    if (JSON.stringify(next) === JSON.stringify(this._focus)) { return; }
    this._focus = next;

    if (this._focusNotifyQueued) { return; }
    this._focusNotifyQueued = true;
    Promise.resolve()
      .then(() => {
        this._focusNotifyQueued = false;
        this.context.dynamicDataSourceManager.notifyPropertyChanged(PROCESS_FOCUS_PROPERTY_ID);
      })
      .catch(() => { this._focusNotifyQueued = false; });
  }

  private get _processList(): string {
    return this.properties.processListTitle || DEFAULT_PROCESS_LIST;
  }

  private get _systemList(): string {
    return this.properties.systemListTitle || DEFAULT_SYSTEM_LIST;
  }

  /** Rebuilt when either list name changes, so its caches never go stale. */
  private _getService(): ProcessDetailsService {
    const key = `${this._processList}|${this._systemList}`;
    if (!this._service || this._serviceKey !== key) {
      this._service = new ProcessDetailsService(
        this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl,
        this._processList,
        this._systemList
      );
      this._serviceKey = key;
    }
    return this._service;
  }

  /** The code to display, as published by the connected hierarchy. */
  private _currentCode(): string | undefined {
    const selection = this.properties.selection
      ? this.properties.selection.tryGetValue()
      : undefined;
    return selection && selection.code ? selection.code : undefined;
  }

  /**
   * Reconciles render state with the current selection, starting a fetch if the
   * code changed. Deliberately does not render: it is called *from* render(),
   * and rendering here would paint the panel twice per selection.
   */
  private _syncToSelection(): void {
    const code = this._currentCode();
    if (code === this._code) { return; }
    this._beginLoad(code);
  }

  private _beginLoad(code: string | undefined): void {
    this._code = code;
    this._step = null;
    this._error = undefined;
    this._notFound = undefined;
    this._loading = !!code;
    // A new L3 closes whatever L4 was open — the codes belong to the previous
    // step and would otherwise leave a consumer showing a stale L4.
    this._openSubStepCode = null;

    // Bumped even when there is no code, so an in-flight request for the
    // previous selection cannot land after a deselect.
    const seq = ++this._requestSeq;
    this._publishFocus();
    if (!code) { return; }

    this._getService().loadProcess(code)
      .then((step) => {
        if (seq !== this._requestSeq) { return; } // a newer selection won
        this._step = step;
        this._notFound = step ? undefined : code;
        this._loading = false;
        this._publishFocus();
        this.render();
      })
      .catch((err: Error) => {
        if (seq !== this._requestSeq) { return; }
        this._step = null;
        this._error = err && err.message ? err.message : String(err);
        this._loading = false;
        this._publishFocus();
        this.render();
      });
  }

  private _retry(): void {
    this._beginLoad(this._code);
    this.render();
  }

  /**
   * Reports the connection and data state to the console.
   *
   * Deliberately the console rather than a block inside the panel. This panel's
   * height is an input to the sticky logic — both its own pinning and the
   * position of anything stacked beneath it — and a report rendered inside it
   * would add several hundred pixels to exactly the measurement being
   * diagnosed.
   */
  private _logDiagnostics(): void {
    if (!this.properties.showDiagnostics) { return; }

    const prop = this.properties.selection;
    let sourceTitle: string | undefined;
    let payload: ISelectedProcessStep | undefined;
    if (prop) {
      payload = prop.tryGetValue();
      // Defensive: the source metadata shape is not worth trusting blindly, and
      // a diagnostics readout that throws is worse than useless.
      try {
        const source = prop.tryGetSource();
        sourceTitle = source
          ? (source.metadata && source.metadata.title) || source.id
          : undefined;
      } catch {
        sourceTitle = undefined;
      }
    }

    const fetchState =
      this._error ? 'error'
        : this._loading ? 'loading'
          : this._notFound ? 'not found'
            : this._step ? 'loaded'
              : 'idle';

    // render() runs on every scroll-driven repaint, so without this guard the
    // console fills with identical lines and the useful ones scroll away.
    const key = [
      sourceTitle, payload && payload.code, this._code,
      this._openSubStepCode, fetchState,
    ].join('|');
    if (key === this._lastDiagKey) { return; }
    this._lastDiagKey = key;

    const d = this._getService().diagnostics;
    console.warn(
      `[Process Details] source ${sourceTitle || '(not connected)'}, ` +
      `selected ${(payload && payload.code) || '(none)'}, ` +
      `showing ${this._code || '(none)'} — ${fetchState}, ` +
      `open L4 ${this._openSubStepCode || '(none)'}` +
      `\n  lists "${this._processList}" / "${this._systemList}", ` +
      `last query returned ${d.lastRowCount} row(s)` +
      (d.unresolvedSystems.length > 0
        ? `\n  no logo resolved for: ${d.unresolvedSystems.join(', ')}`
        : '') +
      (d.lastQuery ? `\n  ${d.lastQuery}` : '')
    );
    if (d.processFields.length > 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const t: any = console;
      if (typeof t.table === 'function') { t.table(d.processFields); }
    }
  }




  public render(): void {
    const sample = this.properties.showSampleData;

    // Kicks off a fetch when the connected code has changed since last paint.
    // Safe to call from render(): it never renders, and returns immediately
    // unless the code actually differs.
    if (!sample) { this._syncToSelection(); }

    const element: React.ReactElement<IProcessDetailsProps> = React.createElement(
      ProcessDetails,
      {
        selectedStep:      sample ? SAMPLE_STEP_WITH_SUBSTEPS : this._step,
        isLoading:         sample ? false : this._loading,
        error:             sample ? undefined : this._error,
        notFoundCode:      sample ? undefined : this._notFound,
        onRetry:           () => { this._retry(); },
        // Controlled from the web part rather than left to the component: the
        // open L4 is part of what this web part publishes, so it cannot live
        // in component state where getPropertyValue() can't reach it.
        openSubStepCode:   sample ? undefined : this._openSubStepCode,
        onSubStepToggle:   (code: string | null) => {
          this._openSubStepCode = code;
          this._publishFocus();
          this.render();
        },
        title:             this.properties.title,
        headerIcon:        this.properties.headerIcon,
        showBreadcrumb:    this.properties.showBreadcrumb !== false,
        singleOpenSubStep: this.properties.singleOpenSubStep !== false,
        emptyIcon:         this.properties.emptyIcon,
        emptyHeading:      this.properties.emptyHeading,
        emptyBody:         this.properties.emptyBody,
      }
    );
    ReactDom.render(element, this.domElement);

    this._logDiagnostics();

    // After the render, so the measurement sees the panel's new height —
    // selecting a step with many L4 rows can flip it from top- to
    // bottom-anchored.
    if (this._sticky) { this._sticky.apply(); }
  }

  protected onThemeChanged(_currentTheme: IReadonlyTheme | undefined): void { /* no-op */ }

  protected onDispose(): void {
    this.context.dynamicDataProvider.unregisterAvailableSourcesChanged(this._onAvailableSourcesChanged);
    // Before unmounting, so the host DOM we borrowed is handed back untouched
    // even if this web part is removed from the page.
    if (this._sticky) { this._sticky.dispose(); }
    // Not properties.selection.dispose() — the framework owns it, because it is
    // declared in propertiesMetadata.
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:
              'Shows details for the process step selected in a Process Hierarchy web part. ' +
              'If the hierarchy is deleted and re-added, reconnect it below.',
          },
          groups: [
            {
              groupName: 'Connection',
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: 'Selected process step',
                  fields: [
                    PropertyPaneDynamicField('selection', {
                      label: 'Process step',
                      // Unset, this drills two levels into the connected
                      // property's own value and offers to bind to one of its
                      // individual members (e.g. just ISelectedProcessStep's
                      // "code" or "label") instead of the whole object. This
                      // panel always wants the whole object, so that drill-down
                      // is switched off — see the matching note in
                      // RasciWebPart.ts, which has the same connection shape.
                      propertyValueDepth: 0,
                    }),
                  ],
                  sharedConfiguration: {
                    depth: DynamicDataSharedDepth.Property,
                    source: {
                      sourcesLabel: 'Connect to a process hierarchy',
                      // Restricts the dropdown to Vertical Flow web parts. If it
                      // ever shows empty, comment this out to confirm the source
                      // is registering before suspecting the GUID.
                      filters: { componentId: Guid.parse(VERTICAL_FLOW_COMPONENT_ID) },
                    },
                    property: {
                      filters: { propertyId: SELECTED_STEP_PROPERTY_ID },
                    },
                  },
                }),
              ],
            },
            {
              groupName: 'Lists',
              groupFields: this._listsFetchFailed
                ? [
                  PropertyPaneTextField('processListTitle', {
                    label: 'Process details list',
                    description: 'Couldn\'t read this site\'s lists automatically — type the list name.',
                    placeholder: DEFAULT_PROCESS_LIST,
                  }),
                  PropertyPaneTextField('systemListTitle', {
                    label: 'System master list',
                    description: 'Couldn\'t read this site\'s lists automatically — type the list name.',
                    placeholder: DEFAULT_SYSTEM_LIST,
                  }),
                ]
                : [
                  PropertyPaneDropdown('processListTitle', {
                    label: 'Process details list',
                    options: this._listDropdownOptions(this.properties.processListTitle, DEFAULT_PROCESS_LIST),
                    selectedKey: this.properties.processListTitle || DEFAULT_PROCESS_LIST,
                    disabled: !this._listOptions,
                  }),
                  PropertyPaneDropdown('systemListTitle', {
                    label: 'System master list',
                    options: this._listDropdownOptions(this.properties.systemListTitle, DEFAULT_SYSTEM_LIST),
                    selectedKey: this.properties.systemListTitle || DEFAULT_SYSTEM_LIST,
                    disabled: !this._listOptions,
                  }),
                ],
            },
            {
              groupName: 'Appearance',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Header title',
                  placeholder: 'Process details',
                }),
                PropertyPaneTextField('headerIcon', {
                  label: 'Header icon',
                  description: 'An emoji shown to the left of the header title.',
                  placeholder: '🧾',
                }),
                PropertyPaneToggle('showBreadcrumb', {
                  label: 'Show step number and breadcrumb',
                }),
                PropertyPaneTextField('emptyIcon', {
                  label: 'Unselected-state icon',
                  description: 'Shown above the message when nothing is selected yet.',
                  placeholder: '👈',
                }),
                PropertyPaneTextField('emptyHeading', {
                  label: 'Unselected-state heading',
                  placeholder: 'Select a process to view details',
                }),
                PropertyPaneTextField('emptyBody', {
                  label: 'Unselected-state message',
                  multiline: true,
                  resizable: true,
                  placeholder:
                    'Click on any process in the process hierarchy to see its inputs, procedure and outputs here.',
                }),
                PropertyPaneToggle('stickyPanel', {
                  label: 'Keep panel in view while scrolling',
                }),
                PropertyPaneSlider('stickyTopOffset', {
                  label: 'Offset from top of page',
                  // Needs to reach well past 200: on a published page
                  // SharePoint's own sticky chrome (suite bar, site header,
                  // command bar) lives inside the scrolling region and can
                  // occupy ~300px, which the pinned panel has to clear.
                  min: 0,
                  max: 400,
                  step: 5,
                  disabled: this.properties.stickyPanel === false,
                }),
              ],
            },
            {
              groupName: 'Troubleshooting',
              groupFields: [
                PropertyPaneToggle('showDiagnostics', {
                  label: 'Show diagnostics',
                }),
                PropertyPaneLabel('diagnosticsHelp', {
                  text:
                    'Adds a report under the panel showing the connection, the resolved '
                    + 'list columns and the last query, and logs every pin decision to the '
                    + 'browser console with the ancestor chain it measured. Leave off for '
                    + 'readers — this is for diagnosing a panel that pins in the wrong place.',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
