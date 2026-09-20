/** Candy brand artwork: sidebar mark and wordmark occupants. */

/** Presentation options shared by the Candy mark and wordmark. */
export interface CandyBrandProps {
  /** Height in px; width follows the artwork's aspect ratio. Defaults to 24. */
  size?: number
  /** Extra class for layout placement. */
  className?: string
}

/**
 * Render the Candy mark: a glossy swirl candy, filled with brand color rather than `currentColor`
 * so the mark keeps its identity independent of the surrounding text color and theme.
 * @param props.size - height in px (default 24; width matches, the artwork is square).
 * @param props.className - extra class for layout placement.
 * @returns the mark svg (aria-hidden decorative brand art).
 */
export function CandyBrandMark({ size = 24, className }: CandyBrandProps) {
  return (
    <svg
      width={size}
      height={size}
      className={className}
      viewBox="0 0 256 256"
      fill="none"
      aria-hidden="true"
    >
      <defs>
        <radialGradient id="dsh-candy-mark-body" cx="0.35" cy="0.32" r="0.75">
          <stop offset="0" stopColor="#FFE1EE" />
          <stop offset="0.45" stopColor="#FF8FC0" />
          <stop offset="1" stopColor="#FF4F97" />
        </radialGradient>
      </defs>
      <circle cx="128" cy="128" r="112" fill="url(#dsh-candy-mark-body)" />
      <path
        d="M128 128 m0 -78 a78 78 0 0 1 55 134 a52 52 0 0 1 -91 -11 a33 33 0 0 1 45 -39 a16 16 0 0 1 14 26"
        fill="none"
        stroke="#FFF3F8"
        strokeWidth={17}
        strokeLinecap="round"
      />
    </svg>
  )
}

/** Design width of {@link CandyWordmarkGlyphs} at its 24-unit design height. */
const WORDMARK_DESIGN_WIDTH = 92

/**
 * Render "Candy" as hand-drawn geometric letterforms rather than text, so the wordmark stays
 * brand art (immune to locale translation) instead of hard-coded UI copy.
 * @returns the letterform strokes (`currentColor`, no text node).
 */
function CandyWordmarkGlyphs() {
  return (
    <g stroke="currentColor" strokeWidth={3} strokeLinecap="round" strokeLinejoin="round" fill="none">
      {/* C */}
      <path d="M15 6 A10 10 0 1 0 15 18" />
      {/* A */}
      <path d="M27 2 L20 22 M27 2 L34 22 M22.8 14 L31.2 14" />
      {/* N */}
      <path d="M38 22 L38 2 L52 22 L52 2" />
      {/* D */}
      <path d="M56 2 L56 22 M56 2 A7 10 0 0 1 56 22" />
      {/* Y */}
      <path d="M74 2 L81 12 L81 22 M88 2 L81 12" />
    </g>
  )
}

/**
 * Render the Candy wordmark, pairing the mark with the product name drawn as brand art.
 * @param props.size - mark height in px (default 24); the name scales to match.
 * @param props.className - extra class for layout placement.
 * @param props.includeMark - whether to include the leading mark; defaults to true.
 * @returns the wordmark (aria-hidden decorative brand art).
 */
export function CandyBrandName({ size = 24, className, includeMark = true }: CandyBrandProps & { includeMark?: boolean }) {
  return (
    <span
      className={className}
      aria-hidden="true"
      style={{ display: 'inline-flex', alignItems: 'center', gap: size / 4, lineHeight: 1 }}
    >
      {includeMark ? <CandyBrandMark size={size} /> : null}
      <svg
        width={(size * WORDMARK_DESIGN_WIDTH) / 24}
        height={size}
        viewBox={`0 0 ${WORDMARK_DESIGN_WIDTH} 24`}
        fill="none"
        aria-hidden="true"
      >
        <CandyWordmarkGlyphs />
      </svg>
    </span>
  )
}
