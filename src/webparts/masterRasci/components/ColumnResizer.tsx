import * as React from 'react';
import styles from './MasterRasci.module.scss';

export const MIN_ACTIVITY_WIDTH = 200;
export const MAX_ACTIVITY_WIDTH = 720;
const KEY_STEP = 16;

const clamp = (w: number): number =>
  Math.max(MIN_ACTIVITY_WIDTH, Math.min(MAX_ACTIVITY_WIDTH, Math.round(w)));

interface IColumnResizerProps {
  /** Called with the new width in px, or undefined to go back to the default. */
  onChange: (width: number | undefined) => void;
}

/**
 * A drag handle for the cell it is placed in (which must be positioned). Works by
 * pointer (mouse, pen or touch) and from the keyboard with the arrow keys;
 * double-click or Escape restores the default width.
 */
const ColumnResizer: React.FC<IColumnResizerProps> = ({ onChange }) => {
  const ref = React.useRef<HTMLDivElement>(null);
  const drag = React.useRef<{ x: number; width: number } | undefined>(undefined);
  const [active, setActive] = React.useState(false);

  /** The cell's current rendered width, correct even while it is still on its default. */
  const currentWidth = (): number =>
    ref.current && ref.current.parentElement ? ref.current.parentElement.offsetWidth : 300;

  return (
    <div
      ref={ref}
      className={active ? styles.resizer + ' ' + styles.resizerActive : styles.resizer}
      role="separator"
      aria-orientation="vertical"
      aria-label="Resize the Process column. Use the left and right arrow keys, or double-click to reset."
      aria-valuemin={MIN_ACTIVITY_WIDTH}
      aria-valuemax={MAX_ACTIVITY_WIDTH}
      aria-valuenow={Math.round(currentWidth())}
      tabIndex={0}
      onPointerDown={(e) => {
        e.preventDefault();
        e.currentTarget.setPointerCapture(e.pointerId);
        drag.current = { x: e.clientX, width: currentWidth() };
        setActive(true);
      }}
      onPointerMove={(e) => {
        if (drag.current) { onChange(clamp(drag.current.width + e.clientX - drag.current.x)); }
      }}
      onPointerUp={() => { drag.current = undefined; setActive(false); }}
      onPointerCancel={() => { drag.current = undefined; setActive(false); }}
      onDoubleClick={() => onChange(undefined)}
      onKeyDown={(e) => {
        if (e.key === 'ArrowLeft') { e.preventDefault(); onChange(clamp(currentWidth() - KEY_STEP)); }
        else if (e.key === 'ArrowRight') { e.preventDefault(); onChange(clamp(currentWidth() + KEY_STEP)); }
        else if (e.key === 'Escape') { onChange(undefined); }
      }}
    />
  );
};

export default ColumnResizer;
