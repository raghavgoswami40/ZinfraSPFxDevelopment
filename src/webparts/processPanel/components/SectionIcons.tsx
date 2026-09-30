import * as React from 'react';

/**
 * Inline SVG icons: the Purpose / Inputs / Outputs labels and the two section
 * header logos. Inline rather than an icon font so nothing has to load before
 * they render. They draw in currentColor, so they always match the text they
 * sit beside.
 */

const Svg: React.FC<{ children: React.ReactNode; strokeWidth?: number }> = ({
  children, strokeWidth = 2.2,
}) => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={strokeWidth}
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
    focusable="false"
  >
    {children}
  </svg>
);

/** A target: what the process is aiming at. */
export const PurposeIcon: React.FC = () => (
  <Svg>
    <circle cx="12" cy="12" r="9.5" />
    <circle cx="12" cy="12" r="5.5" />
    <circle cx="12" cy="12" r="1.5" fill="currentColor" />
  </Svg>
);

/** An arrow dropping into a tray: something coming in. */
export const InputsIcon: React.FC = () => (
  <Svg>
    <path d="M12 3v12" />
    <path d="M7 10l5 5 5-5" />
    <path d="M4 17v3h16v-3" />
  </Svg>
);

/** The mirror of InputsIcon: an arrow leaving the tray. */
export const OutputsIcon: React.FC = () => (
  <Svg>
    <path d="M12 15V3" />
    <path d="M7 8l5-5 5 5" />
    <path d="M4 17v3h16v-3" />
  </Svg>
);

/* Header logos. A finer stroke than the label icons above, to sit inside a
   tile the way the reference logos do. */

/** A page with text lines: the process's written details. */
export const ProcessDetailsLogo: React.FC = () => (
  <Svg strokeWidth={1.7}>
    <path d="M6 3h8l4 4v14H6z" />
    <path d="M14 3v4h4" />
    <path d="M9 12h6" />
    <path d="M9 16h6" />
  </Svg>
);

/** Two people: who is responsible, accountable and so on. */
export const RasciLogo: React.FC = () => (
  <Svg strokeWidth={1.7}>
    <circle cx="9" cy="8" r="3" />
    <path d="M3.5 20c0-3 2.5-5 5.5-5s5.5 2 5.5 5" />
    <circle cx="17" cy="9" r="2.5" />
    <path d="M16 14.5c3 0 5 1.8 5 4.5" />
  </Svg>
);
