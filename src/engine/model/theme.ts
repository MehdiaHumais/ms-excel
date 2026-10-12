/**
 * Workbook themes (Page Layout > Themes / Colors / Fonts).
 *
 * A theme is stored on the workbook so charts, shapes and theme-referencing
 * cell styles all resolve against the same palette.
 */

export interface ThemeDef {
  id: string;
  name: string;
  /** Six accent colors, accent1..accent6, in Excel's order. */
  accents: string[];
  headingFont: string;
  bodyFont: string;
}

const OFFICE_ACCENTS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];

export const DEFAULT_THEME: ThemeDef = {
  id: 'office',
  name: 'Office',
  accents: OFFICE_ACCENTS,
  headingFont: 'Calibri Light',
  bodyFont: 'Calibri'
};

/** Theme gallery entries (Page Layout > Themes). */
export const THEMES: ThemeDef[] = [
  DEFAULT_THEME,
  { id: 'office-classic', name: 'Office Classic', accents: ['#5b9bd5', '#ed7d31', '#a5a5a5', '#ffc000', '#4472c4', '#70ad47'], headingFont: 'Calibri', bodyFont: 'Calibri' },
  { id: 'grayscale', name: 'Grayscale', accents: ['#d0cece', '#a5a5a5', '#7f7f7f', '#595959', '#3b3b3b', '#1f1f1f'], headingFont: 'Calibri Light', bodyFont: 'Calibri' },
  { id: 'blue-warm', name: 'Blue Warm', accents: ['#4a66ac', '#629dd1', '#297fd5', '#7f8fa9', '#5a2a82', '#9d90a0'], headingFont: 'Calibri Light', bodyFont: 'Calibri' },
  { id: 'blue', name: 'Blue', accents: ['#0f6fc6', '#009dd9', '#0bd0d9', '#10cf9b', '#7cca62', '#a5c249'], headingFont: 'Calibri Light', bodyFont: 'Calibri' },
  { id: 'green', name: 'Green', accents: ['#548235', '#70ad47', '#a9d18e', '#c6e0b4', '#375623', '#76933c'], headingFont: 'Calibri Light', bodyFont: 'Calibri' },
  { id: 'orange', name: 'Orange', accents: ['#c55a11', '#ed7d31', '#f4b183', '#f8cbad', '#833c0c', '#a85b2b'], headingFont: 'Calibri Light', bodyFont: 'Calibri' },
  { id: 'violet', name: 'Violet', accents: ['#5b2d8e', '#8e63c7', '#b3a1d9', '#cdb9e6', '#3b1f5e', '#6b4a99'], headingFont: 'Calibri Light', bodyFont: 'Calibri' }
];

/** Color-only variants (Page Layout > Colors); accents swap, fonts stay. */
export const THEME_COLOR_SETS: { name: string; accents: string[] }[] = [
  { name: 'Office', accents: OFFICE_ACCENTS },
  { name: 'Office 2007-2010', accents: ['#4f81bd', '#c0504d', '#9bbb59', '#8064a2', '#4bacc6', '#f79646'] },
  { name: 'Grayscale', accents: ['#d0cece', '#a5a5a5', '#7f7f7f', '#595959', '#3b3b3b', '#1f1f1f'] },
  { name: 'Blue Warm', accents: ['#4a66ac', '#629dd1', '#297fd5', '#7f8fa9', '#5a2a82', '#9d90a0'] },
  { name: 'Blue', accents: ['#0f6fc6', '#009dd9', '#0bd0d9', '#10cf9b', '#7cca62', '#a5c249'] },
  { name: 'Blue II', accents: ['#1c4e80', '#2e75b6', '#9dc3e6', '#bdd7ee', '#5b9bd5', '#41719c'] },
  { name: 'Green', accents: ['#548235', '#70ad47', '#a9d18e', '#c6e0b4', '#375623', '#76933c'] },
  { name: 'Green Yellow', accents: ['#4f6228', '#98d020', '#d2ed8b', '#f3f8d0', '#77933c', '#9bbb59'] },
  { name: 'Yellow', accents: ['#bf9000', '#ffc000', '#ffd966', '#fff2cc', '#7f6000', '#a6a600'] },
  { name: 'Orange', accents: ['#c55a11', '#ed7d31', '#f4b183', '#f8cbad', '#833c0c', '#a85b2b'] },
  { name: 'Red', accents: ['#a50e0e', '#e02020', '#f4795b', '#f9cb9c', '#a5192e', '#e03c31'] },
  { name: 'Red Orange', accents: ['#c00000', '#ff5a1f', '#f8a175', '#fbd5b5', '#a52523', '#e46c0a'] },
  { name: 'Violet', accents: ['#5b2d8e', '#8e63c7', '#b3a1d9', '#cdb9e6', '#3b1f5e', '#6b4a99'] },
  { name: 'Violet II', accents: ['#403151', '#745481', '#a08abb', '#c3b0d5', '#251432', '#564a73'] },
  { name: 'Median', accents: ['#94b6d2', '#dd8047', '#a5ab81', '#d8b25c', '#7ba79d', '#968c8c'] },
  { name: 'Paper', accents: ['#a5b592', '#f3b4d4', '#eaa5b0', '#e9c3d9', '#a5c1d9', '#b2c4d8'] }
];

/** Font-pair variants (Page Layout > Fonts). */
export const THEME_FONT_SETS: { name: string; heading: string; body: string }[] = [
  { name: 'Office', heading: 'Calibri Light', body: 'Calibri' },
  { name: 'Calibri', heading: 'Calibri', body: 'Calibri' },
  { name: 'Arial', heading: 'Arial', body: 'Arial' },
  { name: 'Corbel', heading: 'Corbel', body: 'Corbel' },
  { name: 'Candara', heading: 'Candara', body: 'Candara' },
  { name: 'Franklin Gothic', heading: 'Franklin Gothic Medium', body: 'Franklin Gothic Book' },
  { name: 'Georgia', heading: 'Georgia', body: 'Georgia' },
  { name: 'Times New Roman', heading: 'Times New Roman', body: 'Times New Roman' },
  { name: 'Trebuchet MS', heading: 'Trebuchet MS', body: 'Trebuchet MS' },
  { name: 'Tw Cen MT', heading: 'Tw Cen MT', body: 'Tw Cen MT' },
  { name: 'Verdana', heading: 'Verdana', body: 'Verdana' }
];

export function getTheme(theme: ThemeDef | undefined): ThemeDef {
  return theme ?? DEFAULT_THEME;
}

/** Resolve a 0-based accent index (0..5) to a hex color. */
export function accent(theme: ThemeDef | undefined, index: number): string {
  const t = getTheme(theme);
  return t.accents[((index % t.accents.length) + t.accents.length) % t.accents.length];
}

/**
 * Apply the workbook theme to the UI chrome.
 *
 * Page Layout > Themes / Colors / Fonts must visibly change the app the way
 * Excel's own theme does, so the accent channels and fonts are published as
 * custom properties that index.css maps onto --accent / --font-body.
 */
export function applyThemeToChrome(theme: ThemeDef): void {
  const el = document.documentElement;
  const accent1 = theme.accents[0];
  const lum = luminance(accent1);
  el.style.setProperty('--wb-accent', rgbChannels(accent1));
  // Both modes get their own soft tint; index.css picks one via .dark.
  // Light mode tints toward white (toward black when the accent is pale);
  // dark mode tints toward black, lifting very dark accents so they stay visible.
  el.style.setProperty('--wb-accent-soft-light', rgbChannels(mixHex(accent1, lum > 0.55 ? -0.72 : 0.86)));
  el.style.setProperty('--wb-accent-soft-dark', rgbChannels(mixHex(accent1, lum < 0.08 ? 0.25 : -0.75)));
  el.style.setProperty('--wb-font-body', `"${theme.bodyFont}", Calibri, sans-serif`);
  el.style.setProperty('--wb-font-heading', `"${theme.headingFont}", Calibri, sans-serif`);
  // Grid header tints follow the accent so the sheet picks up the new theme.
  el.style.setProperty('--wb-header', rgbChannels(mixHex(accent1, 0.93)));
  el.style.setProperty('--wb-header-dark', rgbChannels(mixHex(accent1, lum > 0.55 ? 0.35 : -0.78)));
}

function parseHex(hex: string): [number, number, number] {
  const h = hex.trim().replace('#', '');
  const full = h.length === 3 ? h[0] + h[0] + h[1] + h[1] + h[2] + h[2] : h.padEnd(6, '0').slice(0, 6);
  return [parseInt(full.slice(0, 2), 16), parseInt(full.slice(2, 4), 16), parseInt(full.slice(4, 6), 16)];
}

/** `rgb(var(--accent))` needs space-separated channels, so expose the triplet. */
export function rgbChannels(hex: string): string {
  return parseHex(hex).join(' ');
}

/** Blend a hex color toward white (amount 0..1) or black when negative. */
export function mixHex(hex: string, amount: number): string {
  const [r, g, b] = parseHex(hex);
  const target = amount >= 0 ? 255 : 0;
  const t = Math.abs(amount);
  const ch = (c: number) => Math.round(c + (target - c) * t);
  return `#${[ch(r), ch(g), ch(b)].map((c) => c.toString(16).padStart(2, '0')).join('')}`;
}

/** Perceived luminance 0..1, used to keep accents legible in both modes. */
export function luminance(hex: string): number {
  const [r, g, b] = parseHex(hex).map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}