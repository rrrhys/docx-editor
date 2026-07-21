/**
 * Tests for ruler indent-marker hit-test priority.
 *
 * Regression: with a negative paragraph indent, the blue indent markers sit
 * inside the grey margin zone (which is a full-height margin drag handle).
 * Clicks near a marker glyph must start an indent drag, not a margin drag.
 */

import { describe, test, expect } from 'bun:test';
import {
  resolveIndentMarkerHit,
  MARKER_GRAB_RADIUS,
  type RulerMarkerCandidate,
} from './rulerHitTest';

const RULER_HEIGHT = 22;

// Left margin at 96px; paragraph indent -566 twips ≈ -37px → markers at ~59px,
// well inside the grey margin zone (0..96px).
const markersAtNegativeIndent: RulerMarkerCandidate[] = [
  { type: 'firstLineIndent', positionPx: 59, zone: 'top' },
  { type: 'leftIndent', positionPx: 59, zone: 'bottom' },
];

describe('resolveIndentMarkerHit', () => {
  test('click exactly on marker inside margin zone grabs the indent marker', () => {
    expect(resolveIndentMarkerHit(59, 18, RULER_HEIGHT, markersAtNegativeIndent)).toBe(
      'leftIndent'
    );
  });

  test('click a few px off the glyph still grabs the marker (not the margin)', () => {
    expect(
      resolveIndentMarkerHit(59 + MARKER_GRAB_RADIUS, 18, RULER_HEIGHT, markersAtNegativeIndent)
    ).toBe('leftIndent');
    expect(
      resolveIndentMarkerHit(59 - MARKER_GRAB_RADIUS, 18, RULER_HEIGHT, markersAtNegativeIndent)
    ).toBe('leftIndent');
  });

  test('vertical zone disambiguates stacked markers at the same x', () => {
    // Top half → first-line indent (▼), bottom half → left indent (▲)
    expect(resolveIndentMarkerHit(59, 4, RULER_HEIGHT, markersAtNegativeIndent)).toBe(
      'firstLineIndent'
    );
    expect(resolveIndentMarkerHit(59, 20, RULER_HEIGHT, markersAtNegativeIndent)).toBe(
      'leftIndent'
    );
  });

  test('click far from any marker falls through to margin drag (null)', () => {
    expect(
      resolveIndentMarkerHit(59 + MARKER_GRAB_RADIUS + 1, 18, RULER_HEIGHT, markersAtNegativeIndent)
    ).toBe(null);
    expect(resolveIndentMarkerHit(10, 10, RULER_HEIGHT, markersAtNegativeIndent)).toBe(null);
  });

  test('wrong-half click near a lone marker still grabs it over the margin', () => {
    const onlyLeftIndent: RulerMarkerCandidate[] = [
      { type: 'leftIndent', positionPx: 59, zone: 'bottom' },
    ];
    // Click in the top half, but only the bottom-half marker is nearby:
    // grabbing the indent marker beats silently moving the page margin.
    expect(resolveIndentMarkerHit(59, 4, RULER_HEIGHT, onlyLeftIndent)).toBe('leftIndent');
  });

  test('right indent marker inside right margin zone wins over margin', () => {
    const markers: RulerMarkerCandidate[] = [
      { type: 'rightIndent', positionPx: 700, zone: 'top' },
    ];
    expect(resolveIndentMarkerHit(702, 5, RULER_HEIGHT, markers)).toBe('rightIndent');
  });

  test('nearest marker wins when two markers are within range', () => {
    const markers: RulerMarkerCandidate[] = [
      { type: 'firstLineIndent', positionPx: 55, zone: 'top' },
      { type: 'rightIndent', positionPx: 62, zone: 'top' },
    ];
    expect(resolveIndentMarkerHit(56, 4, RULER_HEIGHT, markers)).toBe('firstLineIndent');
    expect(resolveIndentMarkerHit(61, 4, RULER_HEIGHT, markers)).toBe('rightIndent');
  });
});
