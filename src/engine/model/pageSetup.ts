/** Page layout defaults, shared by the Page Layout ribbon and the Page Setup dialog. */
export const DEFAULT_MARGINS = {
  top: 0.75,
  bottom: 0.75,
  left: 0.7,
  right: 0.7,
  header: 0.3,
  footer: 0.3
};

/** Paper sizes in inches, matching Excel's "Size" gallery. */
export const PAPER_SIZES: { id: string; label: string; wIn: number; hIn: number }[] = [
  { id: 'letter', label: 'Letter', wIn: 8.5, hIn: 11 },
  { id: 'legal', label: 'Legal', wIn: 8.5, hIn: 14 },
  { id: 'executive', label: 'Executive', wIn: 7.5, hIn: 10 },
  { id: 'a4', label: 'A4', wIn: 8.27, hIn: 11.69 },
  { id: 'a5', label: 'A5', wIn: 5.83, hIn: 8.27 },
  { id: 'tabloid', label: 'Tabloid', wIn: 11, hIn: 17 },
  { id: 'envelope', label: 'Envelope #10', wIn: 4.12, hIn: 9.5 }
];
