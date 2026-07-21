/**
 * Pure hit-testing logic for the horizontal ruler's indent markers.
 *
 * Kept free of React/DOM imports so it can be unit-tested directly.
 */

/** Triangle half-width in px (matches the glyph in HorizontalRuler). */
export const TRI_SIZE = 5;
/** Extra grabbable px on each side of a triangle glyph. */
export const MARKER_HIT_PAD = 3;
/** Px within which an indent marker wins over the margin zone. */
export const MARKER_GRAB_RADIUS = TRI_SIZE + MARKER_HIT_PAD + 1;

/** Candidate indent marker for hit-testing (position in px from ruler left edge). */
export interface RulerMarkerCandidate {
  type: 'firstLineIndent' | 'leftIndent' | 'rightIndent';
  positionPx: number;
  /** Vertical zone the marker glyph occupies: 'top' (▼) or 'bottom' (▲). */
  zone: 'top' | 'bottom';
}

/**
 * Resolve which indent marker (if any) should win a mousedown at (x, y).
 *
 * When a paragraph has a negative indent, the blue indent markers sit inside
 * the grey margin zone, which is itself a full-height margin drag handle.
 * The tiny triangle glyphs are easy to miss by a pixel or two, in which case
 * the margin zone would start a margin drag instead. This gives indent
 * markers priority within `grabRadius` px of their glyph, preferring a
 * marker whose vertical zone matches the click, then the nearest.
 *
 * Returns the winning marker type, or null (caller falls back to margin drag).
 */
export function resolveIndentMarkerHit(
  x: number,
  y: number,
  rulerHeight: number,
  candidates: RulerMarkerCandidate[],
  grabRadius: number = MARKER_GRAB_RADIUS
): RulerMarkerCandidate['type'] | null {
  const clickZone: 'top' | 'bottom' = y < rulerHeight / 2 ? 'top' : 'bottom';

  let best: RulerMarkerCandidate | null = null;
  let bestScore = Infinity;

  for (const c of candidates) {
    const dist = Math.abs(x - c.positionPx);
    if (dist > grabRadius) continue;
    // Prefer markers in the clicked vertical half; penalize the other half
    // (but still allow it — better to grab the wrong-half indent marker than
    // to silently move the page margin).
    const score = dist + (c.zone === clickZone ? 0 : grabRadius + 1);
    if (score < bestScore) {
      bestScore = score;
      best = c;
    }
  }

  return best ? best.type : null;
}
