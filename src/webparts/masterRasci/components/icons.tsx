import * as React from 'react';

/** Small inline icons. Decorative: every use sits next to visible text or inside a
 *  button that carries its own accessible name, so they are hidden from readers. */
const base = {
  width: 16,
  height: 16,
  viewBox: '0 0 16 16',
  fill: 'none',
  stroke: 'currentColor',
  strokeWidth: 1.6,
  strokeLinecap: 'round' as const,
  strokeLinejoin: 'round' as const,
  'aria-hidden': true,
  focusable: false,
};

export const SearchIcon: React.FC = () => (
  <svg {...base}><circle cx="7" cy="7" r="4.5" /><path d="M10.5 10.5 14 14" /></svg>
);

export const CloseIcon: React.FC = () => (
  <svg {...base} strokeWidth={1.8}><path d="M4 4l8 8M12 4l-8 8" /></svg>
);

export const ListIcon: React.FC = () => (
  <svg {...base}><path d="M5.5 4h8M5.5 8h8M5.5 12h8" /><circle cx="2.5" cy="4" r=".6" /><circle cx="2.5" cy="8" r=".6" /><circle cx="2.5" cy="12" r=".6" /></svg>
);

export const MatrixIcon: React.FC = () => (
  <svg {...base}><rect x="2" y="2" width="12" height="12" rx="2" /><path d="M2 6.5h12M2 10h12M6.5 2v12M10 2v12" /></svg>
);

/** Process group: stacked layers. */
export const LayersIcon: React.FC = () => (
  <svg {...base}><path d="M8 2l6 3-6 3-6-3 6-3z" /><path d="M2 8l6 3 6-3" /><path d="M2 11l6 3 6-3" /></svg>
);

/** View as role: a person. */
export const UserIcon: React.FC = () => (
  <svg {...base}><circle cx="8" cy="5.5" r="2.5" /><path d="M3 14c0-2.8 2.2-4.5 5-4.5s5 1.7 5 4.5" /></svg>
);

/** Involvement filter: a funnel. */
export const FunnelIcon: React.FC = () => (
  <svg {...base}><path d="M2 3h12l-4.5 5.5V13l-3-1.5V8.5L2 3z" /></svg>
);

export const SunIcon: React.FC = () => (
  <svg {...base}><circle cx="8" cy="8" r="2.8" /><path d="M8 1.5v1.6M8 12.9v1.6M1.5 8h1.6M12.9 8h1.6M3.4 3.4l1.1 1.1M11.5 11.5l1.1 1.1M3.4 12.6l1.1-1.1M11.5 4.5l1.1-1.1" /></svg>
);

export const MoonIcon: React.FC = () => (
  <svg {...base}><path d="M13.4 9.6A5.6 5.6 0 0 1 6.4 2.6a5.6 5.6 0 1 0 7 7z" /></svg>
);

export const CheckIcon: React.FC = () => (
  <svg {...base} strokeWidth={1.8}><path d="M3.5 8.5l3 3 6-7" /></svg>
);

/** Points down; rotate it with CSS for the collapsed state. */
export const ChevronIcon: React.FC = () => (
  <svg {...base} strokeWidth={1.8}><path d="M4 6l4 4 4-4" /></svg>
);

/** Two chevrons converging: fold every group up. */
export const CollapseAllIcon: React.FC = () => (
  <svg {...base}><path d="M4 3l4 3.5L12 3" /><path d="M4 13l4-3.5L12 13" /></svg>
);

/** Two chevrons diverging: open every group. */
export const ExpandAllIcon: React.FC = () => (
  <svg {...base}><path d="M4 6.5L8 3l4 3.5" /><path d="M4 9.5L8 13l4-3.5" /></svg>
);
