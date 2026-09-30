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
import { StickyPanelController } from '../../shared/stickyPanel';

import ProcessPanel from './components/ProcessPanel';
import {
  IProcessStep, IRasciRecord, RoleKey, EmptyReason, ROLE_ORDER, ROLE_META,
} from './components/IProcessPanelProps';
import { SAMPLE_STEP_WITH_SUBSTEPS } from './components/processDetailsData';
import { resolveColors, isValidHex, DEFAULT_ROLE_COLORS } from './components/roleColors';
import { ProcessDetailsService } from './services/ProcessDetailsService';
import { RasciService } from './services/RasciService';
import {
  ISelectedProcessStep, SELECTED_STEP_PROPERTY_ID, VERTICAL_FLOW_COMPONENT_ID,
} from '../../shared/dynamicData';

export interface IProcessPanelWebPartProps {
  /** Named `selection` rather than `selectedStep` so it cannot be confused with
   *  the component prop of that name. */
  selection: DynamicProperty<ISelectedProcessStep>;

  // ── Lists ─────────────────────────────────────────────────────────────────
  processListTitle: string;
  systemListTitle: string;
  rasciListTitle: string;

  // ── Process Details section ──────────────────────────────────────────────
  pdTitle: string;
  showBreadcrumb: boolean;

  // The two below are deliberately not exposed in the property pane — they are
  // fixed behaviour, not authoring choices. They still take effect, from the
  // manifest defaults; re-add a PropertyPaneToggle for either to make it
  // configurable again.
  /** Accordion behaviour for L4 rows. Fixed on. */
  singleOpenSubStep: boolean;
  /** Renders the worked example instead of live data. Fixed off. */
  showSampleData: boolean;

  pdEmptyIcon: string;
  pdEmptyHeading: string;
  pdEmptyBody: string;

  // ── RASCI section ─────────────────────────────────────────────────────────
  rasciTitle: string;
  rasciEmptyIcon: string;
  /** Shown for 'awaitingL4' instead of rasciEmptyIcon. */
  rasciEmptyIconL4: string;
  rasciEmptyHeadingL3: string;
  rasciEmptyBodyL3: string;
  rasciEmptyHeadingL4: string;
  rasciEmptyBodyL4: string;

  colorResponsible: string;
  colorAccountable: string;
  colorSupports: string;
  colorConsulted: string;
  colorInformed: string;

  // ── Panel-wide behaviour ──────────────────────────────────────────────────
  /** Keeps the whole panel on screen while the page scrolls. */
  stickyPanel: boolean;
  /** Gap left above the pinned panel, to clear SharePoint's own sticky chrome. */
  stickyTopOffset: number;
  /** Console + in-panel report of field resolution and panel pinning. Off by
   *  default, and exposed in the property pane under Troubleshooting. */
  showDiagnostics: boolean;
}

const DEFAULT_PROCESS_LIST = 'Process Details';
const DEFAULT_SYSTEM_LIST = 'Systems Master';
const DEFAULT_RASCI_LIST = 'RASCI';

const DEFAULT_STICKY_TOP = 90;

/** Property-bag key holding each role's colour, so the pane and the render
 *  path agree on the mapping without either hard-coding the other's names. */
const COLOR_PROP: Record<RoleKey, keyof IProcessPanelWebPartProps> = {
  responsible: 'colorResponsible',
  accountable: 'colorAccountable',
  supports:    'colorSupports',
  consulted:   'colorConsulted',
  informed:    'colorInformed',
};

/** What a RASCI lookup should show for the currently focused process, and why
 *  there may be nothing to show. Computed straight from Process Details'
 *  render state — the two sections used to be separate web parts linked by a
 *  published Dynamic Data property; now they are one instance, so this is
 *  just a private method rather than a publish/subscribe round trip. */
interface IRasciTarget {
  code?: string;
  /** The title to display instead of the RASCI list's own — Process Details'
   *  title for the focused L3/L4, so the two lists can never visibly disagree.
   *  Absent when Process Details itself found nothing, in which case the
   *  RASCI list's own Title is used instead. */
  title?: string;
  reason: EmptyReason;
  waiting: boolean;
}

export default class ProcessPanelWebPart extends BaseClientSideWebPart<IProcessPanelWebPartProps> {

  // ── Process Details render state ─────────────────────────────────────────
  private _pdService: ProcessDetailsService | undefined;
  private _pdServiceKey = '';
  private _code: string | undefined = undefined;
  private _step: IProcessStep | null = null;
  private _loading = false;
  private _error: string | undefined = undefined;
  private _notFound: string | undefined = undefined;
  private _requestSeq = 0;
  /** Which L4 row is expanded. null means every row closed, which has to be
   *  told apart from "no L3 selected" — RASCI needs the open L4, not just the
   *  selected L3. */
  private _openSubStepCode: string | null = null;

  // ── RASCI render state ────────────────────────────────────────────────────
  private _rasciService: RasciService | undefined;
  private _rasciServiceKey = '';
  private _rasciCode: string | undefined = undefined;
  private _rasciRecord: IRasciRecord | null = null;
  private _rasciLoading = false;
  private _rasciError: string | undefined = undefined;
  private _rasciNotFound: string | undefined = undefined;
  private _rasciRequestSeq = 0;

  /** Last diagnostics state logged, so repaints do not spam the console. */
  private _lastDiagKey = '';

  /** Class-property arrow, not a method: unregisterAvailableSourcesChanged
   *  matches on function reference, so .bind() at each site would never
   *  unregister. */
  private _onAvailableSourcesChanged = (): void => { this.render(); };

  /** Pins the whole panel (both sections) to the viewport while the page
   *  scrolls. Built in onInit(), not as a field initializer — this.domElement
   *  is not guaranteed to exist yet at construction time. */
  private _sticky: StickyPanelController | undefined = undefined;

  // ── List picker for the property pane ────────────────────────────────────
  // One fetch covers all three list pickers (Process Details, Systems Master,
  // RASCI) — all are plain generic lists on the same site, so a single query
  // for BaseTemplate 100 lists serves every dropdown.
  private _listOptions: IPropertyPaneDropdownOption[] | undefined = undefined;
  private _listsPromise: Promise<IPropertyPaneDropdownOption[]> | undefined = undefined;
  private _listsFetchFailed = false;

  private _loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    if (this._listsPromise) { return this._listsPromise; }

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
          '[Process Panel] Could not read this site\'s lists for the property pane pickers ' +
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
   * Options for one of the three list dropdowns. Always includes the currently
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
      diagnosticsEnabled: () => !!this.properties.showDiagnostics,
      logLabel: 'Process Panel',
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

  private get _processList(): string {
    return this.properties.processListTitle || DEFAULT_PROCESS_LIST;
  }

  private get _systemList(): string {
    return this.properties.systemListTitle || DEFAULT_SYSTEM_LIST;
  }

  private get _rasciList(): string {
    return this.properties.rasciListTitle || DEFAULT_RASCI_LIST;
  }

  /** Rebuilt when either list name changes, so its caches never go stale. */
  private _getPdService(): ProcessDetailsService {
    const key = `${this._processList}|${this._systemList}`;
    if (!this._pdService || this._pdServiceKey !== key) {
      this._pdService = new ProcessDetailsService(
        this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl,
        this._processList,
        this._systemList
      );
      this._pdServiceKey = key;
    }
    return this._pdService;
  }

  /** Rebuilt when the list name changes, so its cache never goes stale. */
  private _getRasciService(): RasciService {
    if (!this._rasciService || this._rasciServiceKey !== this._rasciList) {
      this._rasciService = new RasciService(
        this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl,
        this._rasciList
      );
      this._rasciServiceKey = this._rasciList;
    }
    return this._rasciService;
  }

  /** The code to display, as published by the connected hierarchy. */
  private _currentCode(): string | undefined {
    const selection = this.properties.selection
      ? this.properties.selection.tryGetValue()
      : undefined;
    return selection && selection.code ? selection.code : undefined;
  }

  // ── Process Details fetch cycle ──────────────────────────────────────────

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
    // step and would otherwise leave RASCI showing a stale L4.
    this._openSubStepCode = null;

    // Bumped even when there is no code, so an in-flight request for the
    // previous selection cannot land after a deselect.
    const seq = ++this._requestSeq;
    if (!code) { return; }

    this._getPdService().loadProcess(code)
      .then((step) => {
        if (seq !== this._requestSeq) { return; } // a newer selection won
        this._step = step;
        this._notFound = step ? undefined : code;
        this._loading = false;
        this.render();
      })
      .catch((err: Error) => {
        if (seq !== this._requestSeq) { return; }
        this._step = null;
        this._error = err && err.message ? err.message : String(err);
        this._loading = false;
        this.render();
      });
  }

  private _retryProcessDetails(): void {
    this._beginLoad(this._code);
    this.render();
  }

  // ── RASCI fetch cycle ─────────────────────────────────────────────────────

  /**
   * The code to show a RASCI for, and — when there is none — why.
   *
   * This is the whole of the L3/L4 rule:
   *  - an L3 with sub-processes answers through its L4s, so nothing is shown
   *    until one is opened;
   *  - an L3 without them answers for itself.
   *
   * Reads Process Details' own render state directly rather than re-deriving
   * it from scratch, so the two sections can never disagree about whether the
   * current L3 has sub-processes.
   */
  private _resolveRasciTarget(): IRasciTarget {
    if (!this._code) { return { reason: 'noSelection', waiting: false }; }

    // Process Details is still fetching, so whether this L3 has sub-processes
    // is not yet known. Hold on the loading state rather than briefly showing
    // the L3's own RASCI for a step that turns out to have L4 children.
    if (this._loading) { return { reason: 'noSelection', waiting: true }; }

    const step = this._step;
    const hasSubSteps = !!step && step.subSteps.length > 0;

    if (hasSubSteps) {
      const open = this._openSubStepCode
        ? step!.subSteps.filter((sub) => sub.code === this._openSubStepCode)[0]
        : undefined;
      if (open) { return { code: open.code, title: open.title, reason: 'notFound', waiting: false }; }
      return { reason: 'awaitingL4', waiting: false };
    }

    if (step) { return { code: step.code, title: step.title, reason: 'notFound', waiting: false }; }

    // No sub-processes and no step means Process Details found no row at all
    // for this L3. Its RASCI may still be authored, so fall back to the L3
    // code rather than going blank on the strength of a different list's gap.
    return { code: this._code, reason: 'notFound', waiting: false };
  }

  /**
   * Reconciles RASCI's render state with the current target, starting a fetch
   * if the code changed. Deliberately does not render, for the same reason as
   * _syncToSelection().
   */
  private _syncToRasci(): void {
    const code = this._resolveRasciTarget().code;
    if (code === this._rasciCode) { return; }
    this._beginRasciLoad(code);
  }

  private _beginRasciLoad(code: string | undefined): void {
    this._rasciCode = code;
    this._rasciRecord = null;
    this._rasciError = undefined;
    this._rasciNotFound = undefined;
    this._rasciLoading = !!code;

    const seq = ++this._rasciRequestSeq;
    if (!code) { return; }

    this._getRasciService().loadRasci(code)
      .then((record) => {
        if (seq !== this._rasciRequestSeq) { return; } // a newer selection won
        this._rasciRecord = record;
        this._rasciNotFound = record ? undefined : code;
        this._rasciLoading = false;
        this.render();
      })
      .catch((err: Error) => {
        if (seq !== this._rasciRequestSeq) { return; }
        this._rasciRecord = null;
        this._rasciError = err && err.message ? err.message : String(err);
        this._rasciLoading = false;
        this.render();
      });
  }

  private _retryRasci(): void {
    this._beginRasciLoad(this._rasciCode);
    this.render();
  }

  /**
   * Reports the connection and data state to the console.
   *
   * Deliberately the console rather than a block inside the panel. This panel's
   * height is an input to the sticky logic, and a report rendered inside it
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

    const pdState =
      this._error ? 'error'
        : this._loading ? 'loading'
          : this._notFound ? 'not found'
            : this._step ? 'loaded'
              : 'idle';

    const rasciState =
      this._rasciError ? 'error'
        : this._rasciLoading ? 'loading'
          : this._rasciNotFound ? 'not found'
            : this._rasciRecord ? 'loaded'
              : 'idle';

    // render() runs on every scroll-driven repaint, so without this guard the
    // console fills with identical lines and the useful ones scroll away.
    const key = [
      sourceTitle, payload && payload.code, this._code, this._openSubStepCode, pdState,
      this._rasciCode, rasciState,
    ].join('|');
    if (key === this._lastDiagKey) { return; }
    this._lastDiagKey = key;

    const pd = this._getPdService().diagnostics;
    console.warn(
      `[Process Panel] Process Details — source ${sourceTitle || '(not connected)'}, ` +
      `selected ${(payload && payload.code) || '(none)'}, ` +
      `showing ${this._code || '(none)'} — ${pdState}, ` +
      `open L4 ${this._openSubStepCode || '(none)'}` +
      `\n  lists "${this._processList}" / "${this._systemList}", ` +
      `last query returned ${pd.lastRowCount} row(s)` +
      (pd.unresolvedSystems.length > 0
        ? `\n  no logo resolved for: ${pd.unresolvedSystems.join(', ')}`
        : '') +
      (pd.lastQuery ? `\n  ${pd.lastQuery}` : '')
    );
    if (pd.processFields.length > 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const t: any = console;
      if (typeof t.table === 'function') { t.table(pd.processFields); }
    }

    const rasci = this._getRasciService().diagnostics;
    console.warn(
      `[Process Panel] RASCI — showing ${this._rasciCode || '(none)'} — ${rasciState}` +
      `\n  list "${this._rasciList}", last query returned ${rasci.lastRowCount} row(s)` +
      (rasci.lastQuery ? `\n  ${rasci.lastQuery}` : '')
    );
    if (rasci.fields.length > 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const t: any = console;
      if (typeof t.table === 'function') { t.table(rasci.fields); }
    }
  }

  public render(): void {
    const sample = this.properties.showSampleData;

    // Kicks off a fetch when the connected code has changed since last paint.
    // Safe to call from render(): it never renders, and returns immediately
    // unless the code actually differs.
    if (!sample) { this._syncToSelection(); }
    // Always driven from the real Process Details state, even in sample mode —
    // matches the original two-web-part behaviour, where Process Details only
    // ever published its real fetch state, never the sample.
    this._syncToRasci();

    const target = this._resolveRasciTarget();

    // The step name shown must be the one actually selected — Process Details'
    // own title for the focused L3/L4 — not whatever the RASCI list's
    // separately-authored Title column happens to hold for that code. The two
    // lists are maintained independently, so their Title text can drift; the
    // resolved target is what the reader is looking at, so it wins whenever
    // it is available.
    const rasciRecord: IRasciRecord | null = this._rasciRecord
      ? { ...this._rasciRecord, title: target.title || this._rasciRecord.title }
      : null;

    const element = React.createElement(ProcessPanel, {
      processDetails: {
        selectedStep:      sample ? SAMPLE_STEP_WITH_SUBSTEPS : this._step,
        isLoading:         sample ? false : this._loading,
        error:             sample ? undefined : this._error,
        notFoundCode:      sample ? undefined : this._notFound,
        onRetry:           () => { this._retryProcessDetails(); },
        // Controlled from the web part rather than left to the component: the
        // open L4 feeds RASCI's own target, so it cannot live in component
        // state where _resolveRasciTarget() can't reach it.
        openSubStepCode:   sample ? undefined : this._openSubStepCode,
        onSubStepToggle:   (code: string | null) => {
          this._openSubStepCode = code;
          this.render();
        },
        title:             this.properties.pdTitle,
        showBreadcrumb:    this.properties.showBreadcrumb !== false,
        singleOpenSubStep: this.properties.singleOpenSubStep !== false,
        emptyIcon:         this.properties.pdEmptyIcon,
        emptyHeading:      this.properties.pdEmptyHeading,
        emptyBody:         this.properties.pdEmptyBody,
      },
      rasci: {
        record:         rasciRecord,
        isLoading:      this._rasciLoading || target.waiting,
        error:          this._rasciError,
        emptyReason:    this._rasciNotFound ? 'notFound' : target.reason,
        notFoundCode:   this._rasciNotFound,
        onRetry:        () => { this._retryRasci(); },
        title:          this.properties.rasciTitle,
        colors:         resolveColors({
          responsible: this.properties.colorResponsible,
          accountable: this.properties.colorAccountable,
          supports:    this.properties.colorSupports,
          consulted:   this.properties.colorConsulted,
          informed:    this.properties.colorInformed,
        }),
        emptyIcon:      this.properties.rasciEmptyIcon,
        emptyIconL4:    this.properties.rasciEmptyIconL4,
        emptyHeadingL3: this.properties.rasciEmptyHeadingL3,
        emptyBodyL3:    this.properties.rasciEmptyBodyL3,
        emptyHeadingL4: this.properties.rasciEmptyHeadingL4,
        emptyBodyL4:    this.properties.rasciEmptyBodyL4,
      },
    });
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

  /** Rejects anything that is not a 3- or 6-digit hex colour, so a typo shows
   *  in the pane rather than silently falling back at render time. */
  private _validateColor(value: string): string {
    if (!value || isValidHex(value)) { return ''; }
    return 'Enter a hex colour, for example #009DDC.';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const colorFields = ROLE_ORDER.map((role) =>
      PropertyPaneTextField(COLOR_PROP[role], {
        label: ROLE_META[role].label,
        placeholder: DEFAULT_ROLE_COLORS[role],
        onGetErrorMessage: (value: string) => this._validateColor(value),
        deferredValidationTime: 400,
      })
    );

    const listFields = this._listsFetchFailed
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
        PropertyPaneTextField('rasciListTitle', {
          label: 'RASCI list',
          description: 'Couldn\'t read this site\'s lists automatically — type the list name.',
          placeholder: DEFAULT_RASCI_LIST,
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
        PropertyPaneDropdown('rasciListTitle', {
          label: 'RASCI list',
          options: this._listDropdownOptions(this.properties.rasciListTitle, DEFAULT_RASCI_LIST),
          selectedKey: this.properties.rasciListTitle || DEFAULT_RASCI_LIST,
          disabled: !this._listOptions,
        }),
      ];

    return {
      pages: [
        {
          header: {
            description:
              'Shows details and RASCI for the process step selected in a Process Hierarchy ' +
              'web part. If the hierarchy is deleted and re-added, reconnect it below.',
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
                      // is switched off.
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
              groupFields: listFields,
            },
            {
              groupName: 'Process Details appearance',
              groupFields: [
                PropertyPaneTextField('pdTitle', {
                  label: 'Header title',
                  placeholder: 'Process details',
                }),
                PropertyPaneToggle('showBreadcrumb', {
                  label: 'Show breadcrumb',
                }),
                PropertyPaneTextField('pdEmptyIcon', {
                  label: 'Unselected-state icon',
                  description: 'Shown above the message when nothing is selected yet.',
                  placeholder: '👈',
                }),
                PropertyPaneTextField('pdEmptyHeading', {
                  label: 'Unselected-state heading',
                  placeholder: 'Select a process to view details',
                }),
                PropertyPaneTextField('pdEmptyBody', {
                  label: 'Unselected-state message',
                  multiline: true,
                  resizable: true,
                  placeholder:
                    'Click on any process in the process hierarchy to see its inputs, procedure and outputs here.',
                }),
              ],
            },
            {
              groupName: 'RASCI appearance',
              groupFields: [
                PropertyPaneTextField('rasciTitle', {
                  label: 'Header title',
                  placeholder: 'RASCI',
                }),
                PropertyPaneLabel('colourHelp', {
                  text: 'Hex colours for each role’s heading. The colour is '
                    + 'lightened automatically to stay readable on the dark panel.',
                }),
                ...colorFields,
                PropertyPaneTextField('rasciEmptyIcon', {
                  label: 'Unselected-state icon',
                  description: 'Shown above the message when nothing is focused yet.',
                  placeholder: '👈',
                }),
                PropertyPaneTextField('rasciEmptyIconL4', {
                  label: 'Awaiting-L4 icon',
                  description:
                    'Shown above the awaiting-L4 message. Points up at the process '
                    + 'details section by default, where an L4 step is actually selected.',
                  placeholder: '👆',
                }),
                PropertyPaneTextField('rasciEmptyHeadingL3', {
                  label: 'Unselected-state heading',
                  description: 'Shown when nothing has been selected in the hierarchy.',
                  placeholder: 'Select a process step to view its RASCI',
                }),
                PropertyPaneTextField('rasciEmptyBodyL3', {
                  label: 'Unselected-state message',
                  multiline: true,
                  resizable: true,
                  placeholder:
                    'Click on any process in the process hierarchy to see who is responsible, '
                    + 'accountable, supporting, consulted and informed.',
                }),
                PropertyPaneTextField('rasciEmptyHeadingL4', {
                  label: 'Awaiting-L4 heading',
                  description:
                    'Shown when the selected process has sub-processes, so its RASCI '
                    + 'comes from an L4 rather than from itself.',
                  placeholder: 'Select an L4 process step to view its RASCI',
                }),
                PropertyPaneTextField('rasciEmptyBodyL4', {
                  label: 'Awaiting-L4 message',
                  multiline: true,
                  resizable: true,
                  placeholder:
                    'This process has sub-processes. Click one of the L4 process names in '
                    + 'the process details section to see its RASCI.',
                }),
              ],
            },
            {
              groupName: 'Behaviour',
              groupFields: [
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
                    'Logs the connection, the resolved list columns, the last query for '
                    + 'each section and every pin decision to the browser console, with the '
                    + 'ancestor chain it measured. Leave off for readers — this is for '
                    + 'diagnosing a panel that pins in the wrong place or shows the wrong data.',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
