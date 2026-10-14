/**
 * Font-face resolution.
 *
 * PowerPoint only knows "family + bold/italic flags". Intermediate weights (Light, Medium,
 * SemiBold…) exist in PowerPoint only when the OS installs them as separately named faces,
 * e.g. "Inter SemiBold". The 'family-style' mode emits those names; 'family' keeps the plain
 * family and lets bold/italic flags do what they can.
 *
 * Add brand-specific overrides to FONT_FAMILY_MAP (Figma family → PowerPoint face name).
 */
export const FONT_FAMILY_MAP: Record<string, string> = {
  // 'SF Pro Display': 'SF Pro Display',
  // 'Söhne': 'Soehne',
};

const PLAIN_STYLES = new Set(['', 'regular', 'normal', 'book', 'roman', 'bold', 'text']);

export function resolveFontFace(family: string, style: string, mode: 'family' | 'family-style'): string {
  const base = FONT_FAMILY_MAP[family] ?? family;
  if (mode === 'family') return base;
  const weight = style
    .replace(/\b(italic|oblique)\b/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (PLAIN_STYLES.has(weight.toLowerCase())) return base;
  return `${base} ${weight}`;
}

export function isBold(fontWeight: number, style: string): boolean {
  if (fontWeight) return fontWeight >= 600;
  return /\b(bold|black|heavy|extra ?bold|semi ?bold)\b/i.test(style);
}
