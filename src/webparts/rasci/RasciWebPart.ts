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

import Rasci from './components/Rasci';
import {
  IRasciProps, IRasciRecord, EmptyReason,
  RoleKey, ROLE_ORDER, ROLE_META,
} from './components/IRasciProps';
import { DEFAULT_ROLE_COLORS, isValidHex, resolveColors } from './components/roleColors';
import { RasciService } from './services/RasciService';
import {
  IProcessFocus, PROCESS_FOCUS_PROPERTY_ID, PROCESS_DETAILS_COMPONENT_ID,
} from '../../shared/dynamicData';
import { StickyPanelController, findStackAnchorBottom } from '../../shared/stickyPanel';

export interface IRasciWebPartProps {
  title: string;
  headerIcon: string;

  rasciListTitle: string;

  // ── Unselected states ───────────────────────────────────────────────────
  emptyIcon: string;
  /** Shown for 'awaitingL4' instead of emptyIcon — see IRasciProps. */
  emptyIconL4: string;
  /** Shown when nothing is selected in the hierarchy at all. */
  emptyHeadingL3: string;
  emptyBodyL3: string;
  /** Shown when an L3 with sub-processes is selected but no L4 is open. */
  emptyHeadingL4: string;
  emptyBodyL4: string;

  // ── Tag colours, one per role ───────────────────────────────────────────
  colorResponsible: string;
  colorAccountable: string;
  colorSupports: string;
  colorConsulted: string;
  colorInformed: string;

  /** Keeps the panel on screen while the page scrolls. Stacks directly beneath
   *  a Process Details panel when one precedes it on the page — see
   *  _stackAnchor() — so the pair scrolls together in a fixed top-to-bottom
   *  order rather than each pinning independently to the same offset. */
  stickyPanel: boolean;
  /** Gap left above the pinned panel when there is no Process Details panel to
   *  stack beneath — same convention and same default as Process Details'
   *  own offset, so a lone RASCI behaves identically to it. */
  stickyTopOffset: number;

  /** Console + in-panel report of the connection, field resolution and panel
   *  pinning. Off by default, and exposed in the property pane under
   *  Troubleshooting: when the panel pins somewhere unexpected, the console
   *  readout naming the ancestor it anchored to is the only way to tell why
   *  from outside the page. */
  showDiagnostics: boolean;

  /** Named `focus` rather than `processFocus` so it cannot be confused with
   *  the payload type of that name. */
  focus: DynamicProperty<IProcessFocus>;
}

const DEFAULT_RASCI_LIST = 'RASCI';
const DEFAULT_STICKY_TOP = 90;

/** Property-bag key holding each role's colour, so the pane and the render
 *  path agree on the mapping without either hard-coding the other's names. */
const COLOR_PROP: Record<RoleKey, keyof IRasciWebPartProps> = {
  responsible: 'colorResponsible',
  accountable: 'colorAccountable',
  supports:    'colorSupports',
  consulted:   'colorConsulted',
  informed:    'colorInformed',
};

export default class RasciWebPart extends BaseClientSideWebPart<IRasciWebPartProps> {

  private _service: RasciService | undefined;
  private _serviceKey = '';

  // Render state for the current code.
  private _code: string | undefined = undefined;
  private _record: IRasciRecord | null = null;
  private _loading = false;
  private _error: string | undefined = undefined;
  private _notFound: string | undefined = undefined;

  /** Last diagnostics state logged, so repaints do not spam the console. */
  private _lastDiagKey = '';

  /** Guards against an earlier, slower request overwriting a newer selection. */
  private _requestSeq = 0;

  /** Class-property arrow, not a method: unregisterAvailableSourcesChanged
   *  matches on function reference, so .bind() at each site would never
   *  unregister. */
  private _onAvailableSourcesChanged = (): void => { this.render(); };

  /** Pins the panel to the viewport while the page scrolls. Built in onInit(),
   *  not as a field initializer — this.domElement is not guaranteed to exist
   *  yet at construction time. */
  private _sticky: StickyPanelController | undefined = undefined;

  // ── List picker for the property pane ────────────────────────────────────
  private _listOptions: IPropertyPaneDropdownOption[] | undefined = undefined;
  private _listsPromise: Promise<IPropertyPaneDropdownOption[]> | undefined = undefined;
  private _listsFetchFailed = false;

  private _loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    if (this._listsPromise) { return this._listsPromise; }

    // BaseTemplate 100 is a generic custom list — what the RASCI list is.
    // Filtering to that (and out hidden lists) keeps document libraries,
    // calendars and SharePoint's own system lists out of an author's way.
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
          '[RASCI] Could not read this site\'s lists for the property pane picker ' +
          '— falling back to a plain text field:',
          err
        );
        this._listsFetchFailed = true;
        this._listsPromise = undefined; // let a later pane-open retry
        return [];
      });

    return this._listsPromise;
  }

  /**
   * Always includes the configured value even if the live fetch hasn't returned
   * it — still loading, or the saved list was renamed, deleted or isn't a plain
   * generic list — so the author's existing choice is never silently dropped.
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

  protected onPropertyPaneConfigurationStart(): void {
    if (this._listOptions || this._listsFetchFailed) { return; }
    // _loadLists() swallows its own rejections, so this .catch() only satisfies
    // no-floating-promises — it cannot fire.
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
    // that has never had one, where tryGetValue() would throw.
    if (!this.properties.focus) {
      this.properties.focus = new DynamicProperty<IProcessFocus>(
        this.context.dynamicDataProvider,
        (): void => { this.render(); }
      );
    }

    // DynamicProperty copes with a source appearing late on its own, but not
    // with one disappearing — this is what repaints the empty state when the
    // details panel is removed, instead of leaving a stale RASCI on screen.
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
      // Stacks directly beneath the nearest preceding Process Details panel,
      // pinned or not — this is the whole of "Process Details stays on top,
      // RASCI sits below it". Falls back to no extra anchor (and behaves like
      // a standalone sticky panel) when no such panel is on the page yet.
      extraTopAnchor: () => findStackAnchorBottom('process-details', this.domElement),
      diagnosticsEnabled: () => !!this.properties.showDiagnostics,
      logLabel: 'RASCI',
    });
    this._sticky.attachListeners();

    return super.onInit();
  }

  // Without this the framework will not construct the DynamicProperty for us.
  // 'object', not 'string' — the payload is IProcessFocus.
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'focus': { dynamicPropertyType: 'object' },
    };
  }

  private get _listTitle(): string {
    return this.properties.rasciListTitle || DEFAULT_RASCI_LIST;
  }

  /** Rebuilt when the list name changes, so its caches never go stale. */
  private _getService(): RasciService {
    if (!this._service || this._serviceKey !== this._listTitle) {
      this._service = new RasciService(
        this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl,
        this._listTitle
      );
      this._serviceKey = this._listTitle;
    }
    return this._service;
  }

  private _focus(): IProcessFocus | undefined {
    return this.properties.focus ? this.properties.focus.tryGetValue() : undefined;
  }

  /**
   * The code to show a RASCI for, and — when there is none — why.
   *
   * This is the whole of the L3/L4 rule:
   *  - an L3 with sub-processes answers through its L4s, so nothing is shown
   *    until one is opened;
   *  - an L3 without them answers for itself.
   *
   * The details panel has already worked out which case applies (it is the one
   * holding the sub-step list), so this reads its verdict rather than
   * re-deriving it and risking the two disagreeing.
   */
  private _resolveTarget(): { code?: string; reason: EmptyReason; waiting: boolean } {
    const focus = this._focus();

    if (!focus || !focus.l3Code) {
      return { reason: 'noSelection', waiting: false };
    }

    // The details panel is still fetching, so whether this L3 has sub-processes
    // is not yet known. Hold on the loading state rather than briefly showing
    // the L3's own RASCI for a step that turns out to have L4 children.
    if (focus.pending) {
      return { reason: 'noSelection', waiting: true };
    }

    if (focus.focusCode) {
      return { code: focus.focusCode, reason: 'notFound', waiting: false };
    }

    if (focus.hasSubSteps) {
      return { reason: 'awaitingL4', waiting: false };
    }

    // No focusCode and no sub-processes means the details panel found no row at
    // all for this L3. Its RASCI may still be authored, so fall back to the L3
    // code rather than going blank on the strength of a different list's gap.
    return { code: focus.l3Code, reason: 'notFound', waiting: false };
  }

  /**
   * Reconciles render state with the current focus, starting a fetch if the
   * code changed. Deliberately does not render: it is called *from* render(),
   * and rendering here would paint the panel twice per selection.
   */
  private _syncToFocus(): void {
    const code = this._resolveTarget().code;
    if (code === this._code) { return; }
    this._beginLoad(code);
  }

  private _beginLoad(code: string | undefined): void {
    this._code = code;
    this._record = null;
    this._error = undefined;
    this._notFound = undefined;
    this._loading = !!code;

    // Bumped even when there is no code, so an in-flight request for the
    // previous selection cannot land after a deselect.
    const seq = ++this._requestSeq;
    if (!code) { return; }

    this._getService().loadRasci(code)
      .then((record) => {
        if (seq !== this._requestSeq) { return; } // a newer selection won
        this._record = record;
        this._notFound = record ? undefined : code;
        this._loading = false;
        this.render();
      })
      .catch((err: Error) => {
        if (seq !== this._requestSeq) { return; }
        this._record = null;
        this._error = err && err.message ? err.message : String(err);
        this._loading = false;
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
   * height is an input to the sticky logic — it decides whether the panel fits
   * beneath Process Details and so whether it pins at all — and a report
   * rendered inside it would add several hundred pixels to exactly the
   * measurement being diagnosed.
   */
  private _logDiagnostics(): void {
    if (!this.properties.showDiagnostics) { return; }

    const focus = this._focus();
    const fetchState =
      this._error ? 'error'
        : this._loading ? 'loading'
          : this._notFound ? 'not found'
            : this._record ? 'loaded'
              : 'idle';

    const prop = this.properties.focus;
    let sourceTitle: string | undefined;
    if (prop) {
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

    // render() runs on every scroll-driven repaint, so without this guard the
    // console fills with identical lines and the useful ones scroll away.
    const key = [
      sourceTitle, focus && focus.l3Code, focus && focus.hasSubSteps,
      focus && focus.l4Code, this._code, fetchState,
    ].join('|');
    if (key === this._lastDiagKey) { return; }
    this._lastDiagKey = key;

    const d = this._getService().diagnostics;
    console.warn(
      `[RASCI] source ${sourceTitle || '(not connected)'}, ` +
      `L3 ${(focus && focus.l3Code) || '(none)'}` +
      `${focus && focus.hasSubSteps ? ' (has sub-steps)' : ''}, ` +
      `open L4 ${(focus && focus.l4Code) || '(none)'}, ` +
      `showing ${this._code || '(none)'} — ${fetchState}` +
      `\n  list "${this._listTitle}", last query returned ${d.lastRowCount} row(s)` +
      (d.lastQuery ? `\n  ${d.lastQuery}` : '')
    );
    if (d.fields.length > 0) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const t: any = console;
      if (typeof t.table === 'function') { t.table(d.fields); }
    }
  }

  public render(): void {
    // Kicks off a fetch when the focused code has changed since last paint.
    // Safe to call from render(): it never renders, and returns immediately
    // unless the code actually differs.
    this._syncToFocus();

    const target = this._resolveTarget();

    // The step name shown here must be the one the reader actually selected —
    // Process Details' own title for the focused L3/L4 — not whatever the
    // RASCI list's separately-authored Title column happens to hold for that
    // code. The two lists are maintained independently, so their Title text
    // can drift; the focus payload is what the reader is looking at, so it
    // wins whenever it is available.
    const focus = this._focus();
    const focusTitle = focus ? focus.focusTitle : undefined;
    const record: IRasciRecord | null = this._record
      ? { ...this._record, title: focusTitle || this._record.title }
      : null;

    const element: React.ReactElement<IRasciProps> = React.createElement(
      Rasci,
      {
        record,
        isLoading:      this._loading || target.waiting,
        error:          this._error,
        emptyReason:    this._notFound ? 'notFound' : target.reason,
        notFoundCode:   this._notFound,
        onRetry:        () => { this._retry(); },
        title:          this.properties.title,
        headerIcon:     this.properties.headerIcon,
        colors:         resolveColors({
          responsible: this.properties.colorResponsible,
          accountable: this.properties.colorAccountable,
          supports:    this.properties.colorSupports,
          consulted:   this.properties.colorConsulted,
          informed:    this.properties.colorInformed,
        }),
        emptyIcon:      this.properties.emptyIcon,
        emptyIconL4:    this.properties.emptyIconL4,
        emptyHeadingL3: this.properties.emptyHeadingL3,
        emptyBodyL3:    this.properties.emptyBodyL3,
        emptyHeadingL4: this.properties.emptyHeadingL4,
        emptyBodyL4:    this.properties.emptyBodyL4,
      }
    );
    ReactDom.render(element, this.domElement);

    this._logDiagnostics();

    // After the render, so the measurement sees the panel's new height, and
    // sees Process Details' current pinned/unpinned position for the anchor.
    if (this._sticky) { this._sticky.apply(); }
  }

  protected onThemeChanged(_currentTheme: IReadonlyTheme | undefined): void { /* no-op */ }

  protected onDispose(): void {
    this.context.dynamicDataProvider.unregisterAvailableSourcesChanged(this._onAvailableSourcesChanged);
    // Before unmounting, so the host DOM we borrowed is handed back untouched
    // even if this web part is removed from the page.
    if (this._sticky) { this._sticky.dispose(); }
    // Not properties.focus.dispose() — the framework owns it, because it is
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

    return {
      pages: [
        {
          header: {
            description:
              'Shows the RASCI for the process focused in a Process Details web part. ' +
              'If that panel is deleted and re-added, reconnect it below.',
          },
          groups: [
            {
              groupName: 'Connection',
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  // SPFx renders this as "<label>'s properties" above the field
                  // picker below.
                  label: 'Process Details',
                  fields: [
                    PropertyPaneDynamicField('focus', {
                      label: 'Process',
                      // Default (unset) drills two levels into the connected
                      // property's own value and offers to bind to one of its
                      // individual members instead of the whole object — the
                      // "hasSubSteps" / "pending" entries under this field were
                      // that drill-down, listing whichever of IProcessFocus's
                      // keys happened to be defined at the moment the pane
                      // opened. RASCI always wants the whole IProcessFocus
                      // object, never a single member of it, so that drill-down
                      // is switched off here.
                      propertyValueDepth: 0,
                    }),
                  ],
                  sharedConfiguration: {
                    depth: DynamicDataSharedDepth.Property,
                    source: {
                      sourcesLabel: 'Connect to a process details panel',
                      // Restricts the dropdown to Process Details web parts. If
                      // it ever shows empty, comment this out to confirm the
                      // source is registering before suspecting the GUID.
                      filters: { componentId: Guid.parse(PROCESS_DETAILS_COMPONENT_ID) },
                    },
                    property: {
                      filters: { propertyId: PROCESS_FOCUS_PROPERTY_ID },
                    },
                  },
                }),
              ],
            },
            {
              groupName: 'List',
              groupFields: this._listsFetchFailed
                ? [
                  PropertyPaneTextField('rasciListTitle', {
                    label: 'RASCI list',
                    description: 'Couldn\'t read this site\'s lists automatically — type the list name.',
                    placeholder: DEFAULT_RASCI_LIST,
                  }),
                ]
                : [
                  PropertyPaneDropdown('rasciListTitle', {
                    label: 'RASCI list',
                    options: this._listDropdownOptions(this.properties.rasciListTitle, DEFAULT_RASCI_LIST),
                    selectedKey: this.properties.rasciListTitle || DEFAULT_RASCI_LIST,
                    disabled: !this._listOptions,
                  }),
                ],
            },
            {
              groupName: 'Tag colours',
              groupFields: [
                PropertyPaneLabel('colourHelp', {
                  text: 'Hex colours for each role’s tags. Tag text switches between '
                    + 'black and white automatically, and the role heading is lightened '
                    + 'to stay readable on the dark panel.',
                }),
                ...colorFields,
              ],
            },
            {
              groupName: 'Appearance',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Header title',
                  placeholder: 'RASCI',
                }),
                PropertyPaneTextField('headerIcon', {
                  label: 'Header icon',
                  description: 'An emoji shown to the left of the header title.',
                  placeholder: '👥',
                }),
                PropertyPaneTextField('emptyIcon', {
                  label: 'Unselected-state icon',
                  description: 'Shown above the message when nothing is focused yet.',
                  placeholder: '👈',
                }),
                PropertyPaneTextField('emptyIconL4', {
                  label: 'Awaiting-L4 icon',
                  description:
                    'Shown above the awaiting-L4 message. Points up at the process '
                    + 'details panel by default, where an L4 step is actually selected.',
                  placeholder: '👆',
                }),
                PropertyPaneTextField('emptyHeadingL3', {
                  label: 'Unselected-state heading',
                  description: 'Shown when nothing has been selected in the hierarchy.',
                  placeholder: 'Select a process step to view its RASCI',
                }),
                PropertyPaneTextField('emptyBodyL3', {
                  label: 'Unselected-state message',
                  multiline: true,
                  resizable: true,
                  placeholder:
                    'Click on any process in the process hierarchy to see who is responsible, '
                    + 'accountable, supporting, consulted and informed.',
                }),
                PropertyPaneTextField('emptyHeadingL4', {
                  label: 'Awaiting-L4 heading',
                  description:
                    'Shown when the selected process has sub-processes, so its RASCI '
                    + 'comes from an L4 rather than from itself.',
                  placeholder: 'Select an L4 process step to view its RASCI',
                }),
                PropertyPaneTextField('emptyBodyL4', {
                  label: 'Awaiting-L4 message',
                  multiline: true,
                  resizable: true,
                  placeholder:
                    'This process has sub-processes. Click one of the L4 process names in '
                    + 'the process details panel to see its RASCI.',
                }),
                PropertyPaneToggle('stickyPanel', {
                  label: 'Keep panel in view while scrolling',
                }),
                PropertyPaneLabel('stickyHelp', {
                  text:
                    'The offset below is used only when no Process Details panel precedes '
                    + 'this one on the page — otherwise this panel stacks directly beneath it.',
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
