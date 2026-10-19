import * as React from 'react';

export type PageTheme = 'dark' | 'light';

/** The theme scope of the web part that owns the page's theme: Master RASCI on the RASCI
 *  page, otherwise Process Hierarchy on the home page. Both set data-theme on their root. */
const findScope = (): Element | undefined =>
  document.querySelector('[data-mr-scope][data-theme]') || document.querySelector('[data-theme]') || undefined;

const readTheme = (): PageTheme => {
  const scope = findScope();
  return scope && scope.getAttribute('data-theme') === 'light' ? 'light' : 'dark';
};

/** Follows the theme the reader picks in the Process Hierarchy or Master RASCI web part on the
 *  same page, live. Falls back to dark when neither web part is on the page. */
export const usePageTheme = (enabled: boolean): PageTheme => {
  const [theme, setTheme] = React.useState<PageTheme>(readTheme);

  React.useEffect(() => {
    if (!enabled) return undefined;
    let raf = 0;
    const sync = (): void => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => setTheme(readTheme()));
    };
    sync();
    // The other web part may render after this one, so watch for it appearing as well as
    // for its theme attribute changing.
    const observer = new MutationObserver(sync);
    observer.observe(document.body, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ['data-theme'],
    });
    return () => {
      cancelAnimationFrame(raf);
      observer.disconnect();
    };
  }, [enabled]);

  return enabled ? theme : 'dark';
};
