import { RoleColors, RoleKey } from './IProcessPanelProps';

/**
 * Colour handling for the role headings.
 *
 * The five colours are authored in the property pane, so nothing about them can
 * be assumed: an author may pick a near-black or a near-white. The heading
 * colour is therefore derived from the authored one at render time rather than
 * fixed in the stylesheet.
 */

/** The five defaults, as specified. Also the fallback for a malformed value. */
export const DEFAULT_ROLE_COLORS: RoleColors = {
  responsible: '#016891',
  accountable: '#BFBFBF',
  supports:    '#009DDC',
  consulted:   '#6BA244',
  informed:    '#B3BE35',
};

/** #abc and #aabbcc, with or without the hash. Anything else is not a colour
 *  we can do arithmetic on, so it falls back to the role's default. */
const HEX = /^#?(?:[0-9a-f]{3}|[0-9a-f]{6})$/i;

export const isValidHex = (value: string): boolean => HEX.test((value || '').trim());

interface IRgb { r: number; g: number; b: number }

const parseHex = (value: string): IRgb | undefined => {
  const raw = (value || '').trim().replace(/^#/, '');
  if (!HEX.test(raw)) { return undefined; }
  const full = raw.length === 3
    ? raw.split('').map((c) => c + c).join('')
    : raw;
  return {
    r: parseInt(full.slice(0, 2), 16),
    g: parseInt(full.slice(2, 4), 16),
    b: parseInt(full.slice(4, 6), 16),
  };
};

const toHex = (c: IRgb): string =>
  '#' + [c.r, c.g, c.b]
    .map((v) => {
      const hex = Math.max(0, Math.min(255, Math.round(v))).toString(16);
      return hex.length === 1 ? '0' + hex : hex;
    })
    .join('');

/** WCAG relative luminance — the basis for every contrast decision below. */
const luminance = (c: IRgb): number => {
  const channel = (v: number): number => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(c.r) + 0.7152 * channel(c.g) + 0.0722 * channel(c.b);
};

/**
 * A version of the colour light enough to read as heading text on the dark
 * panel. The hue is kept — the heading and its marker swatch must obviously
 * belong together — and only lightness is raised, by mixing toward white until
 * the luminance clears the target.
 *
 * 0.42 puts the darkest default (#016891, luminance 0.11) at about 8:1 against
 * the panel's black background, comfortably past AA for 13px bold text.
 */
export const onDark = (fill: string): string => {
  const rgb = parseHex(fill);
  if (!rgb) { return '#ffffff'; }

  const TARGET = 0.42;
  if (luminance(rgb) >= TARGET) { return toHex(rgb); }

  // Binary search the mix ratio: cheaper and more accurate than stepping, and
  // luminance is monotonic in the mix amount so it always converges.
  let low = 0;
  let high = 1;
  for (let i = 0; i < 12; i++) {
    const mid = (low + high) / 2;
    const mixed: IRgb = {
      r: rgb.r + (255 - rgb.r) * mid,
      g: rgb.g + (255 - rgb.g) * mid,
      b: rgb.b + (255 - rgb.b) * mid,
    };
    if (luminance(mixed) < TARGET) { low = mid; } else { high = mid; }
  }
  return toHex({
    r: rgb.r + (255 - rgb.r) * high,
    g: rgb.g + (255 - rgb.g) * high,
    b: rgb.b + (255 - rgb.b) * high,
  });
};

/** Falls each role back to its default when the authored value is unusable. */
export const resolveColors = (authored: Partial<RoleColors> | undefined): RoleColors => {
  const out = {} as RoleColors;
  (Object.keys(DEFAULT_ROLE_COLORS) as RoleKey[]).forEach((key) => {
    const value = authored ? authored[key] : undefined;
    out[key] = value && isValidHex(value)
      ? '#' + value.trim().replace(/^#/, '')
      : DEFAULT_ROLE_COLORS[key];
  });
  return out;
};
