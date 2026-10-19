const FONT_LINK_ID = 'ph-outfit-font';

/**
 * Adds the Outfit stylesheet to the page once. It is added at runtime rather
 * than imported in a stylesheet because the SPFx production bundler rejects CSS
 * @import. If the font is blocked, text falls back to Segoe UI.
 */
export const ensureOutfitFont = (): void => {
  if (document.getElementById(FONT_LINK_ID)) return;
  const link = document.createElement('link');
  link.id = FONT_LINK_ID;
  link.rel = 'stylesheet';
  link.href = 'https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700&display=swap';
  document.head.appendChild(link);
};
