import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from './MasterRasci.module.scss';
import { ChevronIcon, CheckIcon } from './icons';
import { scaleOf } from './scale';
import { ThemeContext } from './theme';

export interface IPickerOption {
  value: string;
  label: string;
  /** A small colour badge shown before the label, e.g. the R / A / S / C / I code. */
  badge?: { text: string; bg: string; fg: string };
}
export interface IPickerGroup { name: string; options: IPickerOption[] }

interface IGroupedPickerProps {
  groups: IPickerGroup[];
  /** Selected option value; `allValue` when nothing is filtered. */
  value: string;
  allValue: string;
  /** Text of the "everything" option, e.g. "All roles". */
  allLabel: string;
  icon: React.ReactNode;
  /** Id of the visible field label, so the button is named by label + value. */
  labelId: string;
  /** Accessible name of the open list. */
  listLabel: string;
  /** Widest the open panel may be, in px. */
  maxWidth: number;
  /** Colour of the selected option: the role view uses its own highlight. */
  tone: 'accent' | 'highlight';
  onChange: (value: string) => void;
}

interface IPopoverPos { top: number; left: number; width: number; maxHeight: number; scale: number }

const MARGIN = 12;

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

/**
 * A filter picker for a long, grouped list. A native <select> shows every option
 * in one scroll; this opens a panel instead, with the options as compact chips
 * under a heading per group. Headings are labels (bold, with a rule), options are
 * bordered chips, so it is clear which of the two can be clicked. Rendered through
 * a portal because the results card clips anything that overflows it.
 */
const GroupedPicker: React.FC<IGroupedPickerProps> = ({
  groups, value, allValue, allLabel, icon, labelId, listLabel, maxWidth, tone, onChange,
}) => {
  const { theme } = React.useContext(ThemeContext);
  const [open, setOpen] = React.useState(false);
  const [pos, setPos] = React.useState<IPopoverPos | undefined>(undefined);
  const trigger = React.useRef<HTMLButtonElement>(null);
  const popover = React.useRef<HTMLDivElement>(null);

  const options = (): HTMLElement[] =>
    popover.current
      ? Array.prototype.slice.call(popover.current.querySelectorAll('[role="option"]'))
      : [];

  const close = (restoreFocus: boolean): void => {
    setOpen(false);
    if (restoreFocus && trigger.current) { trigger.current.focus(); }
  };

  const place = React.useCallback((): void => {
    if (!trigger.current) { return; }
    const rect = trigger.current.getBoundingClientRect();
    const k = scaleOf(trigger.current);
    const width = Math.min(maxWidth * k, window.innerWidth - MARGIN * 2);
    const left = Math.max(MARGIN, Math.min(rect.left, window.innerWidth - width - MARGIN));
    const below = window.innerHeight - rect.bottom - MARGIN - 6;
    const above = rect.top - MARGIN - 6;
    // Open upward only when there is clearly more room there.
    if (below < 280 * k && above > below) {
      setPos({ top: Math.max(MARGIN, rect.top - 6 - Math.min(above, 560 * k)), left, width, maxHeight: Math.min(above, 560 * k), scale: k });
    } else {
      setPos({ top: rect.bottom + 6, left, width, maxHeight: Math.max(220 * k, below), scale: k });
    }
  }, [maxWidth]);

  React.useLayoutEffect(() => { if (open) { place(); } }, [open, place]);

  // Move focus into the panel once it is placed: the selected option, else the first.
  React.useEffect(() => {
    if (!open || !pos) { return; }
    const all = options();
    const selected = all.filter((el) => el.getAttribute('aria-selected') === 'true')[0];
    (selected || all[0])?.focus();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, pos === undefined]);

  React.useEffect(() => {
    if (!open) { return; }
    const onDown = (e: MouseEvent): void => {
      const t = e.target as Node;
      if (popover.current?.contains(t) || trigger.current?.contains(t)) { return; }
      setOpen(false);
    };
    const onScroll = (e: Event): void => {
      if (popover.current && e.target instanceof Node && popover.current.contains(e.target)) { return; }
      setOpen(false);
    };
    const onResize = (): void => setOpen(false);
    document.addEventListener('mousedown', onDown);
    window.addEventListener('scroll', onScroll, true);
    window.addEventListener('resize', onResize);
    return () => {
      document.removeEventListener('mousedown', onDown);
      window.removeEventListener('scroll', onScroll, true);
      window.removeEventListener('resize', onResize);
    };
  }, [open]);

  const pick = (next: string): void => {
    onChange(next);
    close(true);
  };

  /** Arrow keys move through the grid; Left/Right step one, Up/Down move a row. */
  const onKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Escape') { e.preventDefault(); close(true); return; }
    if (e.key === 'Tab') { setOpen(false); return; }

    const all = options();
    const idx = all.indexOf(document.activeElement as HTMLElement);
    if (idx < 0) { return; }
    let next: HTMLElement | undefined;

    if (e.key === 'ArrowRight') { next = all[Math.min(all.length - 1, idx + 1)]; }
    else if (e.key === 'ArrowLeft') { next = all[Math.max(0, idx - 1)]; }
    else if (e.key === 'Home') { next = all[0]; }
    else if (e.key === 'End') { next = all[all.length - 1]; }
    else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      const cur = all[idx];
      const dir = e.key === 'ArrowDown' ? 1 : -1;
      const rows = all.filter((el) => (el.offsetTop - cur.offsetTop) * dir > 0);
      if (rows.length > 0) {
        // Nearest row in that direction, then the option closest in x.
        const rowTop = rows.reduce((best, el) =>
          Math.abs(el.offsetTop - cur.offsetTop) < Math.abs(best - cur.offsetTop) ? el.offsetTop : best,
          rows[0].offsetTop);
        next = rows.filter((el) => el.offsetTop === rowTop).reduce((best, el) =>
          Math.abs(el.offsetLeft - cur.offsetLeft) < Math.abs(best.offsetLeft - cur.offsetLeft) ? el : best);
      }
    } else { return; }

    e.preventDefault();
    if (next) { next.focus(); }
  };

  let label = allLabel;
  let selected: IPickerOption | undefined;
  groups.forEach((g) => g.options.forEach((o) => { if (o.value === value) { label = o.label; selected = o; } }));

  const onClass = tone === 'highlight' ? styles.pickOptOn : styles.pickOptOnAccent;

  return (
    <>
      <span className={styles.searchWrap}>
        <span className={styles.searchIcon}>{icon}</span>
        <button
          ref={trigger}
          type="button"
          className={cx(styles.control, styles.pickerTrigger)}
          aria-haspopup="listbox"
          aria-labelledby={`${labelId} ${labelId}-value`}
          aria-expanded={open}
          onClick={() => setOpen((o) => !o)}
        >
          <span className={styles.pickerValue} id={`${labelId}-value`}>
            {selected && selected.badge && (
              <span className={styles.pickBadge} style={{ background: selected.badge.bg, color: selected.badge.fg }}>
                {selected.badge.text}
              </span>
            )}
            <span className={styles.pickerText}>{label}</span>
          </span>
          <span className={cx(styles.pickerChevron, open && styles.pickerChevronOpen)}><ChevronIcon /></span>
        </button>
      </span>

      {open && ReactDom.createPortal(
        <div
          ref={popover}
          className={styles.picker}
          data-theme={theme}
          role="listbox"
          aria-label={listLabel}
          onKeyDown={onKeyDown}
          style={pos
            ? ({ top: pos.top, left: pos.left, width: pos.width, maxHeight: pos.maxHeight, '--u': `${pos.scale}px` } as React.CSSProperties)
            : { visibility: 'hidden', top: 0, left: 0 }}
        >
          <button
            type="button"
            role="option"
            aria-selected={value === allValue}
            className={cx(styles.pickOpt, styles.pickAll, value === allValue && onClass)}
            onClick={() => pick(allValue)}
          >
            <span>{allLabel}</span>
            {value === allValue && <CheckIcon />}
          </button>

          {groups.map((g) => (
            <div key={g.name} role="group" aria-label={g.name} className={styles.pickGroup}>
              {g.name && <div className={styles.pickGroupHead} aria-hidden="true">{g.name}</div>}
              <div className={styles.pickGrid}>
                {g.options.map((o) => (
                  <button
                    key={o.value}
                    type="button"
                    role="option"
                    aria-selected={o.value === value}
                    className={cx(styles.pickOpt, o.value === value && onClass)}
                    onClick={() => pick(o.value)}
                  >
                    <span className={styles.pickLabel}>
                      {o.badge && (
                        <span className={styles.pickBadge} style={{ background: o.badge.bg, color: o.badge.fg }}>
                          {o.badge.text}
                        </span>
                      )}
                      <span>{o.label}</span>
                    </span>
                    {o.value === value && <CheckIcon />}
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>,
        document.body
      )}
    </>
  );
};

export default GroupedPicker;
