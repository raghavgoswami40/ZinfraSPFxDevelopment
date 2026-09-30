/**
 * Pins a web part's single rendered panel to the viewport while the page
 * scrolls, so it stays in view beside much taller content. Used by Process
 * Panel, which renders exactly one panel as the sole child of its domElement.
 *
 * Uses `position: fixed`, not `position: sticky`. Sticky can only stick within
 * the nearest *scrolling* ancestor, and in SharePoint the panel sits inside a
 * section that scrolls independently of the page — so sticky pinned it to the
 * section and it left the screen the moment the page scrolled. Fixed is
 * viewport-relative, which makes nested scroll containers irrelevant. The cost
 * is that position has to be recomputed on scroll, and the host's domElement
 * has to hold the vacated space so the column does not collapse.
 *
 * This used to be shared by two separate web parts — Process Details and
 * RASCI — stacked on top of each other, which needed extra choreography
 * (stackFollower/extraTopAnchor host hooks) so the pair gave way to each other
 * as a unit. Now that they are merged into one Process Panel instance there is
 * only ever one panel to pin, so that choreography is gone; see git history on
 * this file for the two-panel version if it is ever needed again.
 */

const DEFAULT_STICKY_TOP = 90;
/** Breathing room between a pinned panel and the section it sits in. */
export const STICKY_GAP = 12;
/** A pinned panel shorter than this is not worth pinning — it would be a
 *  sliver too small to read, so it stays in normal flow instead. */
const STICKY_MIN_BAND = 200;
/** Below this the section stacks into one column, where pinning would cover
 *  the content beside a panel. */
const STICKY_MIN_WIDTH = 768;
/** How far up the host DOM a walk is allowed to reach. Deliberately small:
 *  everything above this is SharePoint's own chrome, not the page canvas. */
const STICKY_MAX_ANCESTORS = 12;
/**
 * Class fragment identifying the canvas column a web part sits in on a modern
 * SharePoint page. Matched by name rather than measured, because a height test
 * cannot reliably tell a real column from the web part's own ControlZone
 * wrapper: that wrapper is the panel's height plus its own padding, and that
 * padding alone (16px on one panel here, 64px on another carrying a bottom
 * margin) straddles any small slack. Landing on the wrapper makes "the top of
 * the section" mean "the top of this web part", so the panel's ceiling tracks
 * its own position and it stops appearing to stick at all.
 */
const COLUMN_CLASS = 'CanvasSection';
/**
 * Fallback only, for a host that is not the modern canvas — the workbench's own
 * mock, or a custom page. An ancestor taller than the panel by at least this
 * much is treated as the stretched column. Deliberately generous for the
 * wrapper-padding reason above: a wrapper is the panel plus tens of pixels,
 * whereas a column the panel can actually travel within is far taller.
 */
const STICKY_TRAVEL_SLACK = 240;

/**
 * The canvas column `el` sits in, or undefined when it is not on a modern
 * canvas.
 */
const nearestColumn = (el: HTMLElement): HTMLElement | undefined => {
  let cursor: HTMLElement | null = el.parentElement;
  for (let i = 0; i < STICKY_MAX_ANCESTORS && cursor; i++) {
    if ((cursor.className || '').toString().indexOf(COLUMN_CLASS) !== -1) {
      return cursor;
    }
    cursor = cursor.parentElement;
  }
  return undefined;
};

export interface IStickyPanelHost {
  /** The web part's own domElement — reserves the vacated space and anchors
   *  the fixed panel's left/width to its normal-flow slot. */
  domElement: HTMLElement;
  /** `=== false` disables sticky outright (the property pane toggle). */
  stickyEnabled: () => boolean;
  /** True while the page is in edit mode — pinning fights the authoring
   *  canvas, drag handles and property pane, so it is suppressed. */
  isEditMode: () => boolean;
  /** Gap this panel asks for from the top of the viewport or section. */
  topOffset: () => number;
  /** Gate for the console diagnostics dump. Omit to disable it entirely. */
  diagnosticsEnabled?: () => boolean;
  /** Prefix for diagnostic console messages, e.g. "Process Details". */
  logLabel: string;
}

export class StickyPanelController {
  private _resizeRaf: number | undefined = undefined;
  /** Last diagnostic state logged, so scroll frames do not spam the console. */
  private _lastLogKey = '';
  /** Decisions whose ancestor table has already been printed once. */
  private readonly _tabledKeys: string[] = [];

  /** Class-property arrow, not a method — removeEventListener matches on
   *  reference. rAF-throttled because scroll and resize both fire
   *  continuously, and each pass reads layout then writes styles. */
  private _onViewportChange = (): void => {
    if (this._resizeRaf !== undefined) { return; }
    this._resizeRaf = window.requestAnimationFrame(() => {
      this._resizeRaf = undefined;
      this.apply();
    });
  };

  public constructor(private readonly _host: IStickyPanelHost) {}

  /** Registers the scroll/resize listeners. Call once from onInit(). Capture
   *  phase on scroll is essential: the page has nested scroll containers (a
   *  section scrolls independently of the page), and scroll events from a
   *  nested container do not bubble — but they are seen on the way down. */
  public attachListeners(): void {
    document.addEventListener('scroll', this._onViewportChange, true);
    window.addEventListener('resize', this._onViewportChange);
  }

  /** Removes the listeners and returns the panel to normal flow. Call once
   *  from onDispose(), before unmounting, so borrowed host DOM is handed back
   *  untouched even if the web part is removed from the page. */
  public dispose(): void {
    document.removeEventListener('scroll', this._onViewportChange, true);
    window.removeEventListener('resize', this._onViewportChange);
    if (this._resizeRaf !== undefined) {
      window.cancelAnimationFrame(this._resizeRaf);
      this._resizeRaf = undefined;
    }
    this.clear();
  }

  /** The rendered panel — the sole child of the host's domElement. */
  private _panel(): HTMLElement | undefined {
    const first = this._host.domElement.firstElementChild;
    return first instanceof HTMLElement ? first : undefined;
  }

  /** Returns the panel to normal flow and releases the reserved space. */
  public clear(): void {
    const panel = this._panel();
    if (panel) {
      const s = panel.style;
      s.position = '';
      s.top = '';
      s.left = '';
      s.width = '';
      s.zIndex = '';
      // In normal flow the scrolling box clips the panel itself, so the clip
      // applied while it was sliding out has to come off with the rest.
      s.clipPath = '';
      s.maxHeight = '';
      s.overflowY = '';
    }
    this._host.domElement.style.height = '';
  }

  /**
   * The nearest ancestor that genuinely scrolls — the section's own scroll box
   * when it has one. Its visible top edge is what "the top of the section"
   * means on screen, and unlike the scrolling *content* inside it, that edge
   * stays put. Returns undefined when nothing between here and the cap
   * scrolls, in which case the page itself is the scroller.
   */
  private _scrollPort(): HTMLElement | undefined {
    // Deliberately uncapped, unlike the other walks. SharePoint's canvas scroll
    // region sits well over a dozen levels above the web part, so a cap here
    // silently reports "the page scrolls" and the panel then pins to a flat
    // viewport offset instead of to the section. This walk only reads, so
    // going all the way to <body> costs nothing.
    let el: HTMLElement | null = this._host.domElement.parentElement;
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

  /**
   * Vertical bounds of the column the panel sits in, in viewport coordinates.
   * The pinned panel is kept inside these: it never rises above the top of its
   * own section (which would float it over the page header), and never outruns
   * the bottom (which would float it over whatever follows the section).
   *
   * Identified by class first and only measured as a fallback — see
   * COLUMN_CLASS for why a height test alone gets this wrong, and how
   * differently it gets it wrong for two panels on the same page.
   */
  private _columnBounds(panelH: number): { top: number; bottom: number; label: string } {
    const bounds = (el: HTMLElement, label: string): { top: number; bottom: number; label: string } => {
      const r = el.getBoundingClientRect();
      return { top: r.top, bottom: r.bottom, label: `${label} ${el.tagName.toLowerCase()}.${(el.className || '').toString().slice(0, 40)}` };
    };

    const column = nearestColumn(this._host.domElement);
    if (column) { return bounds(column, 'by class'); }

    let el: HTMLElement | null = this._host.domElement.parentElement;
    for (let i = 0; i < STICKY_MAX_ANCESTORS && el; i++) {
      if (el.offsetHeight > panelH + STICKY_TRAVEL_SLACK) {
        return bounds(el, 'by height');
      }
      el = el.parentElement;
    }
    // No bounding column; don't clamp.
    return { top: -Infinity, bottom: Infinity, label: 'none' };
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
    let el: HTMLElement | null = this._host.domElement.parentElement;
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
   * depends entirely on host markup that cannot be seen from here, so when it
   * misbehaves this is the only way to find out which ancestor is responsible.
   */
  private _logSticky(key: string, reason: string, isEditMode: boolean): void {
    if (!this._host.diagnosticsEnabled || !this._host.diagnosticsEnabled()) { return; }

    // This runs on every scroll frame, so it has to be deduplicated — but not
    // on the decision alone. A panel that stays in one state for a whole scroll
    // (an unpinned one, say) would then report once, at whichever position it
    // first reached that state, and stay silent through every later position
    // even as the numbers that explain it change completely. Bucketing the
    // panel's own position re-reports roughly every 50px of travel, which is
    // frequent enough to show how the geometry evolves and sparse enough to
    // stay readable.
    const bucket = Math.round(this._host.domElement.getBoundingClientRect().top / 50);
    const dedupe = `${key}@${bucket}`;
    if (dedupe === this._lastLogKey) { return; }
    this._lastLogKey = dedupe;

    // The ancestor table is the bulky part and only changes when the host DOM
    // does, so it goes out once per decision rather than once per report.
    const withTable = this._tabledKeys.indexOf(key) === -1;
    if (withTable) { this._tabledKeys.push(key); }

    const rows: Array<Record<string, unknown>> = [];
    let el: HTMLElement | null = this._host.domElement;
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
      `[${this._host.logLabel}] pin: ${reason} — panel ${this._host.domElement.offsetHeight}px, ` +
      `viewport ${window.innerWidth}x${window.innerHeight}, ` +
      `mode ${isEditMode ? 'Edit' : 'Read'}` +
      (blocker ? `\n  WARNING: fixed positioning is captured by an ancestor: ${blocker}` : '')
    );
    if (!withTable) { return; }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const t: any = console;
    if (typeof t.table === 'function') { t.table(rows); } else { console.warn(rows); }
  }

  /**
   * Pins the panel to the viewport so it stays in view while the page scrolls.
   * Call after every render, once the DOM reflects the new content — sizing a
   * panel that has just changed height (e.g. gained L4 rows) needs the fresh
   * measurement.
   */
  public apply(): void {
    const panel = this._panel();
    if (!panel) { return; }

    const isEditMode = this._host.isEditMode();

    // `=== false`, not `!`: a web part instance added before this property
    // existed has nothing in its property bag, and undefined should mean "on".
    if (this._host.stickyEnabled() === false) {
      this.clear();
      this._logSticky('off-toggle', 'off (toggle)', isEditMode);
      return;
    }
    if (isEditMode) {
      this.clear();
      this._logSticky('off-edit', 'off (edit mode)', isEditMode);
      return;
    }
    if (window.innerWidth < STICKY_MIN_WIDTH) {
      this.clear();
      this._logSticky('off-narrow', 'off (viewport too narrow)', isEditMode);
      return;
    }

    // The panel's full content height, not its rendered height: while pinned it
    // may be capped and scrolling inside itself, and offsetHeight would then
    // report the cap. scrollHeight is the same as offsetHeight when uncapped.
    const panelH = panel.scrollHeight;
    if (panelH === 0) { return; } // not laid out yet; the next render will catch it

    // Where the panel sits in normal flow. While pinned, the host's domElement
    // still occupies that slot at the panel's height, so this stays correct in
    // both states — which is what makes the pin/unpin test stable in both
    // scroll directions.
    const slot = this._host.domElement.getBoundingClientRect();
    const vh = window.innerHeight;
    const requestedTop = this._host.topOffset();
    const top = typeof requestedTop === 'number' ? requestedTop : DEFAULT_STICKY_TOP;

    const column = this._columnBounds(panelH);
    const port = this._scrollPort();
    const portRect = port ? port.getBoundingClientRect() : undefined;

    // "Top of the section" is the scroll box's own top edge when the section
    // scrolls independently — that edge is stable, whereas the tall content
    // inside it slides away. Only when nothing nested scrolls does the column
    // stand in for it. Either way the offset still applies once that edge has
    // scrolled up past it, so the panel keeps clearing SharePoint's chrome.
    //
    // The port is used whatever its height. It used to be discarded when
    // shorter than the panel, on the theory that such a port was some inner
    // widget — but every scroller found here is an *ancestor* of the panel, so
    // it is the page's own scroll region. Discarding it for a tall panel dropped
    // the top edge back to the column's, which has long since scrolled away, and
    // the panel then pinned above the page header.
    const sectionTop = portRect ? portRect.top : column.top;

    // The panel's own ceiling — the highest point it may pin to.
    const ceiling = Math.max(top, sectionTop + STICKY_GAP);

    // The panel also must not hang below the visible bottom of the section.
    const visibleBottom = Math.min(
      vh - STICKY_GAP,
      portRect ? portRect.bottom - STICKY_GAP : Infinity
    );

    // The room the pinned panel has: from its ceiling down to the visible
    // bottom. A `position: fixed` panel never reveals more of itself as the page
    // scrolls, so one taller than this band would lose its lower rows for good.
    // It is instead capped to the band and scrolls inside itself, which keeps
    // every row reachable and keeps it pinned however tall its content gets.
    const band = visibleBottom - ceiling;
    const capped = panelH > band;
    const shownH = capped ? band : panelH;

    let pinTop: number | undefined;

    // A band too small to be usable (a very short window, or the section
    // scrolled almost out of view) is left in normal flow rather than pinned
    // as a sliver.
    if (band >= STICKY_MIN_BAND && slot.top < ceiling) {
      // Pin as soon as the panel's natural position rises above the ceiling.
      pinTop = ceiling;
    }

    if (pinTop === undefined) {
      this.clear();
      this._logSticky(
        'unpinned',
        `unpinned — slotTop ${Math.round(slot.top)}px vs ceiling ${Math.round(ceiling)}px, ` +
        `visibleBottom ${Math.round(visibleBottom)}px, band ${Math.round(band)}px, ` +
        `scrollPort ${port ? `${port.tagName.toLowerCase()} ${Math.round(port.getBoundingClientRect().height)}px tall` : 'none'}`,
        isEditMode
      );
      return;
    }

    // Ride up with the end of the column rather than floating past it. At the
    // very end of a section the panel has to leave, even if that means going
    // above the section top.
    if (column.bottom - shownH < pinTop) {
      pinTop = column.bottom - shownH;
    }

    // Once it has slid far enough that none of it is left inside the section,
    // there is nothing to show and no reason to hold it fixed — hand it back to
    // the flow so it scrolls away like ordinary content.
    if (pinTop + shownH <= sectionTop) {
      this.clear();
      this._logSticky(
        'pushed-out',
        `slid out of the section — would pin at ` +
        `${Math.round(pinTop)}px, section top ${Math.round(sectionTop)}px`,
        isEditMode
      );
      return;
    }

    // Reserve the vacated space before fixing, or the column collapses and
    // everything below it jumps.
    //
    // The panel's own bottom margin is counted in: it is the panel's trailing
    // breathing room in normal flow, and a fixed element's margin affects
    // nothing, so leaving it out would shorten the page by that much the
    // moment the panel pins and lengthen it again when it releases.
    const marginBottom = parseFloat(window.getComputedStyle(panel).marginBottom) || 0;
    this._host.domElement.style.height = `${panelH + marginBottom}px`;

    const s = panel.style;
    s.position = 'fixed';
    s.top = `${pinTop}px`;
    s.left = `${slot.left}px`;
    s.width = `${slot.width}px`;
    s.zIndex = '10';

    // Too tall for the band: cap it there and let it scroll inside itself. The
    // overflow is on the panel, whose own rounded-corner clipping still applies.
    s.maxHeight = capped ? `${Math.floor(band)}px` : '';
    s.overflowY = capped ? 'auto' : '';

    // Hide whatever has slid above the section's top edge. A fixed panel is not
    // clipped by the scrolling box it came from, so without this the part that
    // has given way would be painted over the page header — the one thing the
    // sliding must not be allowed to do. Only ever set while the panel really
    // has risen past that edge, so an ordinary pinned panel carries no clip.
    const hiddenAbove = sectionTop - pinTop;
    s.clipPath = hiddenAbove > 0 ? `inset(${Math.ceil(hiddenAbove)}px 0 0 0)` : '';

    this._logSticky(
      `pinned-${capped ? 'capped' : 'full'}-${port ? 'port' : 'page'}`,
      `pinned at ${Math.round(pinTop)}px (${capped ? `capped to ${Math.round(band)}px of ${Math.round(panelH)}px, scrolls inside` : 'full height'}) — ` +
      `scrollPort ${port ? `${port.tagName.toLowerCase()}.${(port.className || '').toString().slice(0, 30)} top ${Math.round(port.getBoundingClientRect().top)}px h${Math.round(port.getBoundingClientRect().height)}` : 'none (the page scrolls)'}, ` +
      `column [${column.label}] top ${Math.round(column.top)}px bottom ${Math.round(column.bottom)}px, ` +
      `ceiling ${Math.round(ceiling)}px`,
      isEditMode
    );
  }
}
