import * as React from 'react';
import styles from './PageBanner.module.scss';
import Sparkles from './Sparkles';
import { usePageTheme, PageTheme } from './usePageTheme';

// eslint-disable-next-line @typescript-eslint/no-var-requires
const logoOnDark: string = require('../assets/zinfra-logo.png');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const logoOnLight: string = require('../assets/zinfra-logo-dark.png');

export interface IPageBannerProps {
  title: string;
  /** Fixed height in px. Leave out to let the banner grow around its children. */
  height?: number;
  /** Background in the dark theme. */
  backgroundColor: string;
  showSparkles: boolean;
  /** Switch to the light look when the page's Process Hierarchy / Master RASCI is light. */
  followPageTheme: boolean;
  /** Used by a web part that owns the theme itself (Process Hierarchy). */
  theme?: PageTheme;
  /** Edges line up with the host web part's content margins. */
  embedded?: boolean;
  /** Set false when the host places the logo itself. */
  showLogo?: boolean;
  /** Set false to hide the title row; the toolbar in `children` and the starfield stay. */
  showTitle?: boolean;
  /** Shown under the title and logo, over the starfield. */
  children?: React.ReactNode;
}

/** Matches the light canvas of the Process Hierarchy and Master RASCI web parts. */
const LIGHT_BACKGROUND = '#F4F6F9';

const PageBanner: React.FC<IPageBannerProps> = ({ title, height, backgroundColor, showSparkles, followPageTheme, theme: ownTheme, embedded, showLogo = true, showTitle = true, children }) => {
  const pageTheme = usePageTheme(followPageTheme && !ownTheme);
  const theme = ownTheme || pageTheme;
  const light = theme === 'light';
  const background = light ? LIGHT_BACKGROUND : backgroundColor;

  // The line starts under the second letter, so it begins where the first letter ends.
  const firstLetterRef = React.useRef<HTMLSpanElement>(null);
  const [stripLeft, setStripLeft] = React.useState(0);
  React.useLayoutEffect(() => {
    const measure = (): void => {
      const el = firstLetterRef.current;
      if (el) setStripLeft(el.offsetLeft + el.offsetWidth);
    };
    measure();
    window.addEventListener('resize', measure);
    const fonts = (document as unknown as { fonts?: { ready: Promise<unknown> } }).fonts;
    if (fonts) fonts.ready.then(measure).catch(() => undefined);
    return () => window.removeEventListener('resize', measure);
  }, [title]);

  return (
    <div
      className={`${styles.banner} ${light ? styles.light : ''} ${embedded ? styles.embedded : ''} ${showTitle ? '' : styles.noTitle}`}
      style={{
        height: height || undefined,
        background,
        ['--banner-bg' as string]: background,
        ['--strip-left' as string]: `${stripLeft}px`,
      }}
    >
      {showTitle && <div className={styles.row}>
        <h1 className={styles.title}>
          <span ref={firstLetterRef}>{title.charAt(0)}</span>{title.slice(1)}
        </h1>
        {showLogo && (
          <img
            className={styles.logo}
            src={light ? logoOnLight : logoOnDark}
            alt="Zinfra - Your trusted partner in energy"
          />
        )}
      </div>}
      {children && <div className={styles.below}>{children}</div>}
      {showSparkles && (
        <div className={styles.sparkles} aria-hidden="true">
          <div className={`${styles.glow} ${styles.glowBlue}`} />
          <div className={`${styles.line} ${styles.glowBlue}`} />
          <div className={`${styles.glow} ${styles.glowGreen}`} />
          <div className={`${styles.line} ${styles.glowGreen}`} />
          <Sparkles className={styles.canvas} color={light ? '#0B7BA8' : '#FFFFFF'} />
          <div className={styles.fade} />
        </div>
      )}
    </div>
  );
};

export default PageBanner;
