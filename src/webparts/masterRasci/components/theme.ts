import * as React from 'react';

export type ThemeName = 'dark' | 'light';

const KEY = 'mr-theme';

/** The reader's last choice, or dark (the default) when nothing is stored or storage is blocked. */
export const readStoredTheme = (): ThemeName => {
  try { return window.localStorage.getItem(KEY) === 'light' ? 'light' : 'dark'; }
  catch { return 'dark'; }
};

export const storeTheme = (theme: ThemeName): void => {
  try { window.localStorage.setItem(KEY, theme); }
  catch { /* private window or blocked storage: the choice just will not be remembered */ }
};

export interface IThemeContext {
  theme: ThemeName;
  setTheme: (theme: ThemeName) => void;
}

/** Lets the toggle (in the toolbar) and the dropdown panels (rendered on <body>, outside
 *  the themed wrapper) share one theme without passing it through every component. */
export const ThemeContext = React.createContext<IThemeContext>({
  theme: 'dark',
  setTheme: () => undefined,
});
