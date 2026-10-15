import * as React from 'react';
import styles from './MasterRasci.module.scss';
import { ThemeContext } from './theme';
import { SunIcon, MoonIcon } from './icons';

const cx = (...names: Array<string | false | undefined>): string =>
  names.filter(Boolean).join(' ');

/** Light / dark switch, in the same segmented style as the By process / Matrix toggle. */
const ThemeToggle: React.FC = () => {
  const { theme, setTheme } = React.useContext(ThemeContext);
  return (
    <div className={cx(styles.seg, styles.themeToggle)} role="group" aria-label="Theme">
      <button
        type="button"
        aria-pressed={theme === 'light'}
        aria-label="Light theme"
        title="Light theme"
        className={cx(styles.segBtn, styles.segBtnIcon, theme === 'light' && styles.segBtnOn)}
        onClick={() => setTheme('light')}
      ><SunIcon /></button>
      <button
        type="button"
        aria-pressed={theme === 'dark'}
        aria-label="Dark theme"
        title="Dark theme"
        className={cx(styles.segBtn, styles.segBtnIcon, theme === 'dark' && styles.segBtnOn)}
        onClick={() => setTheme('dark')}
      ><MoonIcon /></button>
    </div>
  );
};

export default ThemeToggle;
