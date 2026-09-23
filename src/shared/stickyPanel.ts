/**
 * Pins a web part's single rendered panel to the viewport while the page
 * scrolls, so it stays in view beside (or, when stacked, below) much taller
 * content. Shared by Process Details and RASCI — both render exactly one
 * panel as the sole child of their domElement, and need the same behaviour.
 *
 * Uses `position: fixed`, not `position: sticky`. Sticky can only stick within
 * the nearest *scrolling* ancestor, and in SharePoint the panel sits inside a
 * section that scrolls independently of the page — so sticky pinned it to the
 * section and it left the screen the moment the page scrolled. Fixed is
 * viewport-relative, which makes nested scroll containers irrelevant. The cost
 * is that position has to be recomputed on scroll, and the host's domElement
 * has to hold the vacated space so the column does not collapse.
 */

const DEFAULT_STICKY_TOP = 90;
/** Breathing room between a pinned panel and whatever edge is anchoring it —
 *  the section it sits in, or (when stacked) the panel above it. */
export const STICKY_GAP = 12;
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
 * canvas. Shared by the column-bounds walk and the stack lookups below, so that
 * "the same column" means the same thing to all of them.
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

/**
 * Whether two panels share a canvas column. Document order alone is not enough
 * to call one panel the other's neighbour: a page can hold the same pair of web
 * parts in a second column, or in a later section, and stacking against one of
 * those would position a panel against something nowhere near it on screen.
 *
 * Two panels outside any recognised column are treated as sharing one, so a
 * non-canvas host keeps whatever behaviour document order gives it.
 */
const sameColumn = (a: HTMLElement, b: HTMLElement): boolean =>
  nearestColumn(a) === nearestColumn(b);

/** Geometry of the panel stacked directly beneath a sticky panel. */
export interface IStackFollower {
  /**
   * Top edge of the follower's reserved slot, in viewport pixels — its host
   * element, not its panel. The panel may be fixed or relatively nudged; the
   * slot is neither, so this moves only with scroll and content.
   */
  slotTop: number;
  /** The follower panel's own natural height. */
  height: number;
}

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
  /**
   * An additional floor for the pin position, in viewport pixels — e.g. the
   * bottom edge of another panel this one must stack beneath. Combined with
   * topOffset via max(), so whichever is more restrictive wins. Returns
   * undefined when there is nothing to stack under.
   */
  extraTopAnchor?: () => number | undefined;
  /**
   * Geometry of the panel stacked directly *beneath* this one, if any — its
   * reserved slot's top edge and its natural height.
   *
   * Drives how far this panel gives way. While the pair fits the visible band
   * it gives way not at all and pins as it would alone. When the pair does not
   * fit, holding the pin would park the follower below the fold for good —
   * scrolling cannot move a fixed panel — so this panel slides up to keep its
   * bottom exactly STICKY_GAP above the follower's slot, and is clipped to its
   * section so none of what has slid past the top edge is painted over the page
   * header.
   *
   * Both values are read from the follower's *slot*, never its rendered panel,
   * so both are pure functions of scroll and content and are unmoved by either
   * panel's pin state. That is what keeps the two controllers from feeding into
   * each other: the dependency runs strictly one way — this panel, then the
   * follower — and settles in a single pass.
   */
  stackFollower?: () => IStackFollower | undefined;
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

    const panelH = panel.offsetHeight;
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

    // The panel's own floor, independent of anything it might be stacked
    // beneath. Kept deliberately separate from `ceiling` below: an anchor can
    // legitimately sit far down an unscrolled page (its bottom edge measured
    // in raw viewport pixels can be a large number simply because nothing has
    // scrolled yet), and folding that into the "is this panel too tall for the
    // screen" test below would misclassify a perfectly short panel as needing
    // bottom-anchored behaviour — which is exactly what let RASCI drift out of
    // its pin logic entirely and scroll straight past Process Details.
    const baseCeiling = Math.max(top, sectionTop + STICKY_GAP);

    // The panel also must not hang below the visible bottom of the section.
    const visibleBottom = Math.min(
      vh - STICKY_GAP,
      portRect ? portRect.bottom - STICKY_GAP : Infinity
    );

    const tallerThanViewport = panelH > visibleBottom - baseCeiling;

    // ── Give way, continuously, to the panel stacked beneath ───────────────
    //
    // Both panels of a stack can only hold their pins while the pair fits the
    // visible band. When it does not, this panel has to surrender its position
    // — otherwise the follower is parked below the fold for good, since page
    // scrolling cannot move a fixed panel.
    //
    // It surrenders by *sliding*, not by letting go. `followerPush` is the pin
    // position at which this panel's bottom sits exactly STICKY_GAP above the
    // follower's own slot, so as the follower rises this panel rides up with
    // it, one pixel per pixel, and the follower travels at its natural scroll
    // rate the whole way. An earlier version released the pin outright the
    // moment the follower reached it; that is continuous while scrolling, but
    // it snaps hundreds of pixels when the *follower* is what changes — select
    // a step, the follower grows, the pair stops fitting, and the panel is
    // dropped from its pin to wherever its flow position had scrolled to. The
    // slide turns that into a shift of just the height the follower gained.
    //
    // Sliding takes this panel above its section's top edge, which on its own
    // would float it over the page header. It is clipped to the section below
    // to prevent exactly that, so the rule the clipping upholds is the visible
    // one — nothing of this panel is ever painted outside its section — rather
    // than a restriction on where it may be positioned.
    //
    // Read from the follower's reserved slot and natural height, never its
    // rendered position: both are pure functions of scroll and content and so
    // are unmoved by either panel's pin state. The dependency stays one-way —
    // this panel, then the follower — and settles in a single pass.
    const follower = this._host.stackFollower && this._host.stackFollower();
    let followerPush: number | undefined;
    if (follower && follower.height > 0) {
      // A tall panel holds its bottom at visibleBottom rather than its top at
      // baseCeiling, so that is where its pinned bottom would actually land.
      // Without this a tall leader would push the follower's ceiling clean off
      // the screen and it could never appear at all.
      const wouldBePinnedBottom = tallerThanViewport ? visibleBottom : baseCeiling + panelH;
      const pairFits = wouldBePinnedBottom + STICKY_GAP + follower.height <= visibleBottom;
      if (!pairFits) {
        followerPush = follower.slotTop - STICKY_GAP - panelH;
      }
    }

    // A second panel stacked beneath another (RASCI under Process Details)
    // must never rise above that other panel's current bottom edge, whether it
    // is pinned or still in normal flow — this is what keeps the pair in a
    // fixed top-to-bottom order regardless of scroll position. Folded in only
    // here, for the actual pin position — never into baseCeiling above.
    const extraAnchor = this._host.extraTopAnchor && this._host.extraTopAnchor();
    const ceiling = Math.max(
      baseCeiling,
      extraAnchor !== undefined ? extraAnchor + STICKY_GAP : -Infinity
    );

    // Whether the anchor is the floor actually in force, rather than merely
    // present. Once the panel above has released and scrolled away, its bottom
    // edge rises past baseCeiling and stops constraining anything — from that
    // point this panel is standalone in every respect that matters, and the
    // clamps below have to treat it that way.
    const anchorBinds = extraAnchor !== undefined && extraAnchor + STICKY_GAP > baseCeiling;

    let pinTop: number | undefined;

    if (anchorBinds && tallerThanViewport) {
      // Stacked and too tall for the band even on its own. Pinning could only
      // hold its top at the anchor floor and leave the rest below the fold,
      // unreachable. Left in normal flow instead, where the page scroll reaches
      // every row of it — it forfeits pinning, which is the lesser loss.
    } else if (!tallerThanViewport) {
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
      // Not pinned — but for a panel stacked beneath another, "not pinned"
      // only means the ordinary hard-pin trigger (slot.top < ceiling) hasn't
      // fired yet, which happens whenever the natural gap below the anchor is
      // already *at least* STICKY_GAP. Left alone, that gap is whatever the
      // page's own layout happens to produce — a static CSS margin can only
      // guess at it, and any margin bigger than what SharePoint's own spacing
      // already adds shows up as exactly the oversized "unscrolled" gap this
      // is fixing. So pull the panel up to sit at exactly `ceiling` instead of
      // leaving the excess in place.
      //
      // Strictly a pull *up*, never a push down. Closing an oversized gap is
      // the whole point; pushing a panel down to meet the anchor would glue a
      // panel that has no room to pin onto the anchor's bottom edge, so the
      // page scroll could never reach its lower half — the very thing the
      // fits-under-anchor test above is avoiding.
      //
      // `position: relative` rather than a margin: it repositions the panel
      // purely visually, without feeding back into the very
      // domElement.getBoundingClientRect() measurement this math is based on.
      if (extraAnchor !== undefined && !tallerThanViewport) {
        // Straight to the ceiling, with no column clamp. For an anchored panel
        // that clamp could only ever ask for a position *above* the anchor's
        // bottom edge — the overlap this whole mechanism exists to prevent —
        // so it is a no-op here at best. (An earlier revision took a
        // Math.max() against the column's end, which on a roomy column is a
        // far larger number and shoved the panel hundreds of pixels down the
        // page: the "RASCI is missing at the top of the page" bug.)
        const delta = Math.round(ceiling - slot.top);

        if (delta <= 0) {
          this._host.domElement.style.height = ''; // normal flow, nothing reserved

          const s = panel.style;
          s.position = delta !== 0 ? 'relative' : '';
          s.top = delta !== 0 ? `${delta}px` : '';
          s.left = '';
          s.width = '';
          s.zIndex = '';

          this._logSticky(
            'anchored-unpinned',
            `not pinned, but holding the anchor gap by nudging ${delta}px — slotTop ` +
            `${Math.round(slot.top)}px, ceiling ${Math.round(ceiling)}px, ` +
            `anchor bottom ${Math.round(extraAnchor)}px`,
            isEditMode
          );
          return;
        }
      }

      this.clear();
      this._logSticky(
        'unpinned',
        `unpinned — slotTop ${Math.round(slot.top)}px vs ceiling ${Math.round(ceiling)}px, ` +
        `visibleBottom ${Math.round(visibleBottom)}px, ` +
        `${tallerThanViewport ? 'bottom' : 'top'}-anchor mode, ` +
        `scrollPort ${port ? `${port.tagName.toLowerCase()} ${Math.round(port.getBoundingClientRect().height)}px tall` : 'none'}`,
        isEditMode
      );
      return;
    }

    // Ride up with the end of the column rather than floating past it. Normally
    // this wins over the ceiling outright: at the very end of a section the
    // panel has to leave, even if that means going above the section top.
    //
    // The exception is a panel whose anchor is still binding — letting the
    // column win there would lift it above the anchor's bottom edge, the
    // overlap this whole mechanism exists to prevent, so the column may only
    // push it *down* to its own end. Gated on `anchorBinds` rather than on the
    // anchor merely existing: once the panel above has released and risen out
    // of the way, this panel needs the ordinary clamp back, or it never leaves
    // the end of the section and floats over whatever follows it.
    if (column.bottom - panelH < pinTop) {
      pinTop = anchorBinds
        ? Math.max(ceiling, column.bottom - panelH)
        : column.bottom - panelH;
    }

    // Give way to the panel beneath, as computed above. Applied last and only
    // downward, so it can shorten this panel's stay at the top but never extend
    // it past any of the bounds already settled.
    if (followerPush !== undefined && followerPush < pinTop) { pinTop = followerPush; }

    // Once it has slid far enough that none of it is left inside the section,
    // there is nothing to show and no reason to hold it fixed — hand it back to
    // the flow so it scrolls away like ordinary content.
    if (pinTop + panelH <= sectionTop) {
      this.clear();
      this._logSticky(
        'pushed-out',
        `slid out of the section for the panel beneath — would pin at ` +
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

    // Hide whatever has slid above the section's top edge. A fixed panel is not
    // clipped by the scrolling box it came from, so without this the part that
    // has given way would be painted over the page header — the one thing the
    // sliding must not be allowed to do. Only ever set while the panel really
    // has risen past that edge, so an ordinary pinned panel carries no clip.
    const hiddenAbove = sectionTop - pinTop;
    s.clipPath = hiddenAbove > 0 ? `inset(${Math.ceil(hiddenAbove)}px 0 0 0)` : '';

    // Nothing caps the panel's height here. A panel only reaches this point when
    // pinning it leaves what is visible of it fully readable: this panel slides
    // out of the way when the pair would not fit, and the anchored-and-too-tall
    // branch declines to pin a follower that cannot fit on its own. So neither
    // needs clipping at the bottom or a scrollbar of its own.

    this._logSticky(
      `pinned-${tallerThanViewport ? 'bottom' : 'top'}-${port ? 'port' : 'page'}`,
      `pinned at ${Math.round(pinTop)}px (${tallerThanViewport ? 'bottom' : 'top'}-anchored) — ` +
      `scrollPort ${port ? `${port.tagName.toLowerCase()}.${(port.className || '').toString().slice(0, 30)} top ${Math.round(port.getBoundingClientRect().top)}px h${Math.round(port.getBoundingClientRect().height)}${portRect ? '' : ' (ignored, too short)'}` : 'none (the page scrolls)'}, ` +
      `column [${column.label}] top ${Math.round(column.top)}px bottom ${Math.round(column.bottom)}px, ` +
      `ceiling ${Math.round(ceiling)}px` +
      (extraAnchor !== undefined ? `, stacked below anchor at ${Math.round(extraAnchor)}px` : ''),
      isEditMode
    );
  }
}

/**
 * Marker attribute set on a panel that another web part may need to stack
 * beneath. A data attribute rather than a shared CSS class because the two
 * panels compile from separate SCSS modules with independently hashed class
 * names — this is the one thing about a panel's identity meant to be read
 * across web part bundles.
 */
export const STICKY_ROLE_ATTR = 'data-sticky-role';

/**
 * Finds the bottom edge (in viewport pixels) of the nearest panel carrying
 * `role` that precedes `before` in the document — i.e. the one this element
 * should stack directly beneath. Works whether that panel is currently pinned
 * (fixed) or still in normal flow, since both report their true on-screen
 * position through getBoundingClientRect(). Returns undefined when no such
 * panel exists yet (not rendered, removed, or this one comes first).
 */
export const findStackAnchorBottom = (role: string, before: HTMLElement): number | undefined => {
  const candidates = document.querySelectorAll<HTMLElement>(`[${STICKY_ROLE_ATTR}="${role}"]`);
  // querySelectorAll returns nodes in document order, so the last one that
  // precedes `before` is the closest preceding match.
  let anchor: HTMLElement | undefined;
  candidates.forEach((el) => {
    // eslint-disable-next-line no-bitwise
    const precedes = el.compareDocumentPosition(before) & Node.DOCUMENT_POSITION_FOLLOWING;
    if (precedes && sameColumn(el, before)) { anchor = el; }
  });
  return anchor ? anchor.getBoundingClientRect().bottom : undefined;
};

/**
 * The mirror of findStackAnchorBottom: the first panel carrying `role` that
 * *follows* `after` in the document and shares its column — the panel stacked
 * directly beneath this one.
 *
 * Reports the follower's reserved slot top and its panel's natural height,
 * never the panel's rendered position. Both are therefore untouched by the
 * follower's own pin state, which is what lets the panel above read them
 * without the two controllers feeding into each other.
 *
 * Returns undefined when there is no such panel, it has not been laid out yet,
 * or it sits in another column — each meaning "nothing stacked beneath", which
 * drops the caller onto its standalone behaviour.
 */
export const findStackFollower = (role: string, after: HTMLElement): IStackFollower | undefined => {
  const candidates = document.querySelectorAll<HTMLElement>(`[${STICKY_ROLE_ATTR}="${role}"]`);
  for (let i = 0; i < candidates.length; i++) {
    const el = candidates[i];
    // eslint-disable-next-line no-bitwise
    const follows = after.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING;
    // First match wins here, unlike the anchor lookup above: document order
    // makes the earliest following panel the nearest one beneath.
    if (follows && el.offsetHeight > 0 && sameColumn(el, after)) {
      const slot = el.parentElement || el;
      return { slotTop: slot.getBoundingClientRect().top, height: el.offsetHeight };
    }
  }
  return undefined;
};
