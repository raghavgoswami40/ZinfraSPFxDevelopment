import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Guid, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
  DynamicDataSharedDepth,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, DynamicProperty } from '@microsoft/sp-component-base';

import ProcessDetails from './components/ProcessDetails';
import {
  IProcessDetailsProps, IProcessStep, IDiagnosticsView,
} from './components/IProcessDetailsProps';
import { SAMPLE_STEP_WITH_SUBSTEPS } from './components/processDetailsData';
import { ProcessDetailsService } from './services/ProcessDetailsService';
import {
  ISelectedProcessStep, SELECTED_STEP_PROPERTY_ID, VERTICAL_FLOW_COMPONENT_ID,
} from '../../shared/dynamicData';

export interface IProcessDetailsWebPartProps {
  title: string;
  headerIcon: string;
  showBreadcrumb: boolean;

  // The three below are deliberately not exposed in the property pane — they
  // are fixed behaviour and developer tooling, not authoring choices. They
  // still take effect, from the manifest defaults; re-add a PropertyPaneToggle
  // for any of them to make it configurable again.
  /** Accordion behaviour for L4 rows. Fixed on. */
  singleOpenSubStep: boolean;
  /** Renders the worked example instead of live data. Fixed off. */
  showSampleData: boolean;
  /** Console + in-panel report of field resolution and panel pinning. Fixed off. */
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
/** Breathing room between the pinned panel and the edge it is anchored to,
 *  applied at whichever edge is doing the anchoring. */
const STICKY_GAP = 12;
/** Below this the section stacks into one column, where pinning would cover
 *  the hierarchy rather than sit beside it. */
const STICKY_MIN_WIDTH = 768;
/** How far up the host DOM the fix-up is allowed to reach. Deliberately small:
 *  everything above this is SharePoint's own chrome, not the page canvas. */
const STICKY_MAX_ANCESTORS = 12;
/** An ancestor taller than the panel by at least this much is the stretched
 *  column — the chain below it now has somewhere to travel, so stop there. */
const STICKY_TRAVEL_SLACK = 40;

export default class ProcessDetailsWebPart extends BaseClientSideWebPart<IProcessDetailsWebPartProps> {

  private _service: ProcessDetailsService | undefined;
  private _serviceKey = '';

  // Render state for the current code.
  private _code: string | undefined = undefined;
  private _step: IProcessStep | null = null;
  private _loading = false;
  private _error: string | undefined = undefined;
  private _notFound: string | undefined = undefined;

  // Guards against an earlier, slower request overwriting a newer selection.
  private _requestSeq = 0;

  /** Class-property arrow, not a method: unregisterAvailableSourcesChanged
   *  matches on function reference, so .bind() at each site would never
   *  unregister. */
  private _onAvailableSourcesChanged = (): void => { this.render(); };

  private _resizeRaf: number | undefined = undefined;
  /** Last diagnostic state logged, so scroll frames do not spam the console. */
  private _lastLogKey = '';

  /** Class-property arrow for the same reason as above — removeEventListener
   *  matches on reference. rAF-throttled because scroll and resize both fire
   *  continuously, and each pass reads layout then writes styles. */
  private _onViewportChange = (): void => {
    if (this._resizeRaf !== undefined) { return; }
    this._resizeRaf = window.requestAnimationFrame(() => {
      this._resizeRaf = undefined;
      this._applySticky();
    });
  };

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

    // A fixed panel has to be repositioned on every scroll. Capture phase is
    // essential: the page has nested scroll containers (the section scrolls
    // independently of the page), and scroll events from a nested container do
    // not bubble — but they are seen on the way down.
    document.addEventListener('scroll', this._onViewportChange, true);
    window.addEventListener('resize', this._onViewportChange);

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

    // Bumped even when there is no code, so an in-flight request for the
    // previous selection cannot land after a deselect.
    const seq = ++this._requestSeq;
    if (!code) { return; }

    this._getService().loadProcess(code)
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

  private _retry(): void {
    this._beginLoad(this._code);
    this.render();
  }

  private _diagnostics(): IDiagnosticsView | undefined {
    if (!this.properties.showDiagnostics) { return undefined; }
    const d = this._getService().diagnostics;

    const prop = this.properties.selection;
    let sourceTitle: string | undefined;
    let payload: ISelectedProcessStep | undefined;
    if (prop) {
      payload = prop.tryGetValue();
      // Defensive: the source metadata shape is not worth trusting blindly, and
      // a diagnostics panel that throws is worse than useless.
      try {
        const source = prop.tryGetSource();
        sourceTitle = source
          ? (source.metadata && source.metadata.title) || source.id
          : undefined;
      } catch {
        sourceTitle = undefined;
      }
    }

    const fetchState: IDiagnosticsView['connection']['fetchState'] =
      this._error ? 'error'
        : this._loading ? 'loading'
          : this._notFound ? 'not found'
            : this._step ? 'loaded'
              : 'idle';

    return {
      connection: {
        hasProperty: !!prop,
        sourceTitle,
        hasValue: !!payload,
        payloadCode: payload ? payload.code : undefined,
        payloadLabel: payload ? payload.label : undefined,
        effectiveCode: this._code,
        fetchState,
      },
      processList: this._processList,
      systemList: this._systemList,
      processFields: d.processFields,
      systemFields: d.systemFields,
      systemLogos: d.systemLogos,
      unresolvedSystems: d.unresolvedSystems,
      lastQuery: d.lastQuery,
      lastRowCount: d.lastRowCount,
    };
  }

  /** The rendered `.panel`, which is the sole child of this.domElement. */
  private _panel(): HTMLElement | undefined {
    const first = this.domElement.firstElementChild;
    return first instanceof HTMLElement ? first : undefined;
  }

  /** Returns the panel to normal flow and releases the reserved space. */
  private _clearSticky(): void {
    const panel = this._panel();
    if (panel) {
      const s = panel.style;
      s.position = '';
      s.top = '';
      s.left = '';
      s.width = '';
      s.zIndex = '';
    }
    this.domElement.style.height = '';
  }

  /**
   * Vertical bounds of the column the panel sits in, in viewport coordinates —
   * the first ancestor meaningfully taller than the panel. The pinned panel is
   * kept inside these: it never rises above the top of its own section (which
   * would float it over the page header), and never outruns the bottom (which
   * would float it over whatever follows the section).
   */
  /**
   * The nearest ancestor that genuinely scrolls — the section's own scroll box
   * when it has one. Its visible top edge is what "the top of the section"
   * means on screen, and unlike the scrolling *content* inside it, that edge
   * stays put. Returns undefined when nothing between here and the cap scrolls,
   * in which case the page itself is the scroller.
   */
  private _scrollPort(): HTMLElement | undefined {
    // Deliberately uncapped, unlike the other walks. SharePoint's canvas scroll
    // region sits well over a dozen levels above the web part, so a cap here
    // silently reports "the page scrolls" and the panel then pins to a flat
    // viewport offset instead of to the section. This walk only reads, so
    // going all the way to <body> costs nothing.
    let el: HTMLElement | null = this.domElement.parentElement;
    while (el && el !== document.body && el !== document.documentElement) {
      const cs = window.getComputedStyle(el);
      if (
        (cs.overflowY === 'auto' || cs.overflowY === 'scroll') &&
        el.scrollHeight > el.clientHeight + 1
      ) {
        return el;
      }
      el = el.parentElement;
    }
    return undefined;
  }

  private _columnBounds(panelH: number): { top: number; bottom: number } {
    let el: HTMLElement | null = this.domElement.parentElement;
    for (let i = 0; i < STICKY_MAX_ANCESTORS && el; i++) {
      if (el.offsetHeight > panelH + STICKY_TRAVEL_SLACK) {
        const r = el.getBoundingClientRect();
        return { top: r.top, bottom: r.bottom };
      }
      el = el.parentElement;
    }
    return { top: -Infinity, bottom: Infinity }; // no bounding column; don't clamp
  }

  /**
   * `position: fixed` needs the viewport as its containing block. Any ancestor
   * with a transform, filter, perspective, will-change or paint containment
   * steals that role, and the panel would then be positioned against *it*
   * instead. Nothing can be done about it from here, but it must be reported —
   * otherwise the symptom is a panel that pins to the wrong place for no
   * visible reason.
   */
  private _fixedBlocker(): string | undefined {
    let el: HTMLElement | null = this.domElement.parentElement;
    for (let i = 0; i < STICKY_MAX_ANCESTORS && el; i++) {
      const cs = window.getComputedStyle(el);
      const contain = (cs as unknown as { contain?: string }).contain || '';
      if (
        cs.transform !== 'none' ||
        cs.filter !== 'none' ||
        cs.perspective !== 'none' ||
        cs.willChange.indexOf('transform') !== -1 ||
        /paint|layout|strict|content/.test(contain)
      ) {
        return `${el.tagName.toLowerCase()}.${(el.className || '').toString().slice(0, 40)} ` +
               `(transform: ${cs.transform}, filter: ${cs.filter}, contain: ${contain || 'none'})`;
      }
      el = el.parentElement;
    }
    return undefined;
  }

  /**
   * Dumps the ancestor chain to the console when diagnostics is on. Positioning
   * depends entirely on host markup we cannot see from here, so when it
   * misbehaves this is the only way to find out which ancestor is responsible.
   */
  private _logSticky(key: string, reason: string): void {
    if (!this.properties.showDiagnostics) { return; }
    // This runs on every scroll frame. Without a state guard it would emit
    // thousands of lines and a console.table per second, which is both
    // unreadable and slow enough to skew what is being measured.
    if (key === this._lastLogKey) { return; }
    this._lastLogKey = key;

    const rows: Array<Record<string, unknown>> = [];
    let el: HTMLElement | null = this.domElement;
    for (let i = 0; i <= STICKY_MAX_ANCESTORS && el && el !== document.documentElement; i++) {
      const cs = window.getComputedStyle(el);
      rows.push({
        level: i === 0 ? 'domElement' : `+${i}`,
        tag: el.tagName.toLowerCase(),
        class: (el.className || '').toString().slice(0, 50),
        offsetHeight: el.offsetHeight,
        overflowY: cs.overflowY,
        scrollH: el.scrollHeight,
        clientH: el.clientHeight,
        // The element sticky WOULD have anchored to — the reason sticky failed
        // here, and the reason fixed does not care.
        scrolls: (cs.overflowY === 'auto' || cs.overflowY === 'scroll') &&
                 el.scrollHeight > el.clientHeight + 1,
        transform: cs.transform === 'none' ? '' : cs.transform,
      });
      el = el.parentElement;
    }

    const blocker = this._fixedBlocker();
    console.warn(
      `[Process Details] pin: ${reason} — panel ${this.domElement.offsetHeight}px, ` +
      `viewport ${window.innerWidth}x${window.innerHeight}, ` +
      `mode ${this.displayMode === DisplayMode.Edit ? 'Edit' : 'Read'}` +
      (blocker ? `\n  WARNING: fixed positioning is captured by an ancestor: ${blocker}` : '')
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const t: any = console;
    if (typeof t.table === 'function') { t.table(rows); } else { console.warn(rows); }
  }

  /**
   * Pins the panel to the viewport so it stays beside the (much taller)
   * hierarchy while the page scrolls.
   *
   * Uses `position: fixed`, not `position: sticky`. Sticky can only stick
   * within the nearest *scrolling* ancestor, and in SharePoint the panel sits
   * inside a section that scrolls independently of the page — so sticky pinned
   * it to the section and it left the screen the moment the page scrolled.
   * Fixed is viewport-relative, which makes nested scroll containers
   * irrelevant. The cost is that position has to be recomputed on scroll, and
   * this.domElement has to hold the vacated space so the column does not
   * collapse.
   */
  private _applySticky(): void {
    const panel = this._panel();
    if (!panel) { return; }

    // `=== false`, not `!`: a web part instance added before this property
    // existed has nothing in its property bag, and undefined should mean "on"
    // — same convention as showBreadcrumb / singleOpenSubStep above.
    if (this.properties.stickyPanel === false) { this._clearSticky(); this._logSticky('off-toggle', 'off (toggle)'); return; }
    // Pinning fights the authoring canvas, the drag handles and the property pane.
    if (this.displayMode === DisplayMode.Edit) { this._clearSticky(); this._logSticky('off-edit', 'off (edit mode)'); return; }
    // Below this the columns stack, so a pinned panel covers the hierarchy.
    if (window.innerWidth < STICKY_MIN_WIDTH) { this._clearSticky(); this._logSticky('off-narrow', 'off (viewport too narrow)'); return; }

    const panelH = panel.offsetHeight;
    if (panelH === 0) { return; } // not laid out yet; the next render will catch it

    // Where the panel sits in normal flow. While pinned, this.domElement still
    // occupies that slot at the panel's height, so this stays correct in both
    // states — which is what makes the pin/unpin test stable in both scroll
    // directions.
    const slot = this.domElement.getBoundingClientRect();
    const vh = window.innerHeight;
    const top = typeof this.properties.stickyTopOffset === 'number'
      ? this.properties.stickyTopOffset
      : DEFAULT_STICKY_TOP;

    const column = this._columnBounds(panelH);
    const port = this._scrollPort();
    let portRect = port ? port.getBoundingClientRect() : undefined;

    // Only treat the scroll port as "the section" if it is actually big enough
    // to hold the panel. A small scrolling ancestor is some inner widget, not
    // the section, and letting it drive the bounds squeezes the usable band
    // below the panel height — which silently prevents pinning entirely.
    if (portRect && portRect.height < panelH + STICKY_GAP) { portRect = undefined; }

    // "Top of the section" is the scroll box's own top edge when the section
    // scrolls independently — that edge is stable, whereas the tall content
    // inside it slides away. Only when nothing nested scrolls does the column
    // stand in for it. Either way the offset still applies once that edge has
    // scrolled up past it, so the panel keeps clearing SharePoint's chrome.
    const sectionTop = portRect ? portRect.top : column.top;
    const ceiling = Math.max(top, sectionTop + STICKY_GAP);

    // The panel also must not hang below the visible bottom of the section.
    const visibleBottom = Math.min(
      vh - STICKY_GAP,
      portRect ? portRect.bottom - STICKY_GAP : Infinity
    );

    const tallerThanViewport = panelH > visibleBottom - ceiling;
    let pinTop: number | undefined;

    if (!tallerThanViewport) {
      // Short panel: pin as soon as its natural position rises above the ceiling.
      if (slot.top < ceiling) { pinTop = ceiling; }
    } else {
      // Tall panel: let it scroll along until its bottom edge reaches the bottom
      // of the visible area, then hold it there — the chosen behaviour.
      if (slot.top + panelH < visibleBottom) {
        pinTop = Math.max(visibleBottom - panelH, ceiling);
      }
    }

    if (pinTop === undefined) {
      this._clearSticky();
      this._logSticky(
        'unpinned',
        `unpinned — slotTop ${Math.round(slot.top)}px vs ceiling ${Math.round(ceiling)}px, ` +
        `visibleBottom ${Math.round(visibleBottom)}px, ` +
        `${tallerThanViewport ? 'bottom' : 'top'}-anchor mode, ` +
        `scrollPort ${port ? `${port.tagName.toLowerCase()} ${Math.round(port.getBoundingClientRect().height)}px tall` : 'none'}`
      );
      return;
    }

    // Ride up with the end of the column rather than floating past it. Applied
    // last so it wins over the ceiling: at the very end of a section the panel
    // has to leave, even if that means going above the section top.
    if (column.bottom - panelH < pinTop) { pinTop = column.bottom - panelH; }

    // Reserve the vacated space before fixing, or the column collapses and
    // everything below it jumps.
    this.domElement.style.height = `${panelH}px`;

    const s = panel.style;
    s.position = 'fixed';
    s.top = `${pinTop}px`;
    s.left = `${slot.left}px`;
    s.width = `${slot.width}px`;
    s.zIndex = '10';

    this._logSticky(
      `pinned-${tallerThanViewport ? 'bottom' : 'top'}-${port ? 'port' : 'page'}`,
      `pinned at ${Math.round(pinTop)}px (${tallerThanViewport ? 'bottom' : 'top'}-anchored) — ` +
      `scrollPort ${port ? `${port.tagName.toLowerCase()}.${(port.className || '').toString().slice(0, 30)} top ${Math.round(port.getBoundingClientRect().top)}px h${Math.round(port.getBoundingClientRect().height)}${portRect ? '' : ' (ignored, too short)'}` : 'none (the page scrolls)'}, ` +
      `columnTop ${Math.round(column.top)}px, ceiling ${Math.round(ceiling)}px`
    );
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
        title:             this.properties.title,
        headerIcon:        this.properties.headerIcon,
        showBreadcrumb:    this.properties.showBreadcrumb !== false,
        singleOpenSubStep: this.properties.singleOpenSubStep !== false,
        diagnostics:       this._diagnostics(),
      }
    );
    ReactDom.render(element, this.domElement);

    // After the render, so the measurement sees the panel's new height —
    // selecting a step with many L4 rows can flip it from top- to
    // bottom-anchored.
    this._applySticky();
  }

  protected onThemeChanged(_currentTheme: IReadonlyTheme | undefined): void { /* no-op */ }

  protected onDispose(): void {
    this.context.dynamicDataProvider.unregisterAvailableSourcesChanged(this._onAvailableSourcesChanged);
    document.removeEventListener('scroll', this._onViewportChange, true);
    window.removeEventListener('resize', this._onViewportChange);
    if (this._resizeRaf !== undefined) {
      window.cancelAnimationFrame(this._resizeRaf);
      this._resizeRaf = undefined;
    }
    // Before unmounting, so the host DOM we borrowed is handed back untouched
    // even if this web part is removed from the page.
    this._clearSticky();
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
              groupFields: [
                PropertyPaneTextField('processListTitle', {
                  label: 'Process details list',
                  placeholder: DEFAULT_PROCESS_LIST,
                }),
                PropertyPaneTextField('systemListTitle', {
                  label: 'System master list',
                  placeholder: DEFAULT_SYSTEM_LIST,
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
          ],
        },
      ],
    };
  }
}
