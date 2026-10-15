/**
 * JS twin of the --u formula in MasterRasci.module.scss, for the few places that
 * work in raw pixels: the dropdown panel (rendered outside the scaled area) and
 * the Process column's drag limits. 1 at normal size, up to 1.5 on very wide screens.
 */
const BASE_WIDTH = 1400;
const SLOPE = 0.00025;
const MAX_SCALE = 1.5;

export const scaleForWidth = (width: number): number =>
  Math.max(1, Math.min(MAX_SCALE, 1 + (width - BASE_WIDTH) * SLOPE));

/** Scale in force around an element: derived from the web part's own width.
 *  Takes null because that is what a React ref holds before it is attached. */
// eslint-disable-next-line @rushstack/no-new-null
export const scaleOf = (el: Element | null | undefined): number => {
  const scope = el ? el.closest('[data-mr-scope]') : null;
  return scope ? scaleForWidth(scope.clientWidth) : 1;
};
