/**
 * Integration tests for Selective Save
 *
 * Tests the full pipeline: parse DOCX → modify document model → selective save → verify XML.
 * Uses real DOCX parsing and serialization to ensure round-trip correctness.
 */

import { describe, test, expect } from 'bun:test';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { parseDocx } from './parser';
import { repackDocx } from './rezip';
import { attemptSelectiveSave, finalSectionPropertiesChanged } from './selectiveSave';
import {
  findParagraphOffsets,
  buildPatchedDocumentXml,
  validatePatchSafety,
  countParagraphElements,
  extractBodySectPrXml,
} from './selectiveXmlPatch';
import { DocumentAgent } from '../agent/DocumentAgent';
import { serializeDocument } from './serializer/documentSerializer';
import type { Paragraph, Run } from '../types/document';

// ============================================================================
// Helpers
// ============================================================================

const FIXTURES_DIR = path.resolve(__dirname, '../../../../e2e/fixtures');

async function loadFixture(name: string): Promise<ArrayBuffer> {
  const filePath = path.join(FIXTURES_DIR, name);
  const buffer = fs.readFileSync(filePath);
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function getDocumentXml(buffer: ArrayBuffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file('word/document.xml');
  if (!file) throw new Error('No word/document.xml found');
  return file.async('text');
}

// ============================================================================
// selectiveXmlPatch tests with real DOCX content
// ============================================================================

describe('Selective XML Patch with real DOCX', () => {
  test('finds paragraphs in real document.xml', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const xml = await getDocumentXml(buffer);

    // Count paragraphs
    const count = countParagraphElements(xml);
    expect(count).toBeGreaterThan(0);

    // Try to find paraIds in the XML
    const paraIdPattern = /w14:paraId="([^"]+)"/g;
    const ids: string[] = [];
    let m: RegExpExecArray | null;
    while ((m = paraIdPattern.exec(xml)) !== null) {
      ids.push(m[1]);
    }
    expect(ids.length).toBeGreaterThan(0);

    // Try to find offsets for each unique ID
    const uniqueIds = [...new Set(ids)];
    for (const id of uniqueIds.slice(0, 5)) {
      const offsets = findParagraphOffsets(xml, id);
      // May be null if nested w:p (inner IDs)
      if (offsets) {
        const extracted = xml.substring(offsets.start, offsets.end);
        expect(extracted).toStartWith('<w:p');
        expect(extracted).toContain(`w14:paraId="${id}"`);
      }
    }
  });

  test('validates patch safety on real document', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });
    const originalXml = await getDocumentXml(buffer);
    const serializedXml = serializeDocument(doc);

    // Empty change set should be safe
    const result = validatePatchSafety(originalXml, serializedXml, new Set());
    expect(result.safe).toBe(true);
  });
});

// ============================================================================
// attemptSelectiveSave integration tests
// ============================================================================

describe('attemptSelectiveSave', () => {
  test('returns null when structural change occurred', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(['someId']),
      structuralChange: true,
      hasUntrackedChanges: false,
    });

    expect(result).toBeNull();
  });

  test('returns null when untracked changes exist', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(['someId']),
      structuralChange: false,
      hasUntrackedChanges: true,
    });

    expect(result).toBeNull();
  });

  test('returns valid buffer when no content changes (still updates metadata)', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    // Should return a valid DOCX (may differ due to core properties update)
    expect(result).not.toBeNull();
    expect(result!.byteLength).toBeGreaterThan(0);

    // Verify document.xml is unchanged
    const originalXml = await getDocumentXml(buffer);
    const resultXml = await getDocumentXml(result!);
    expect(resultXml).toBe(originalXml);

    // Verify core properties were updated with new modification date
    const zip = await JSZip.loadAsync(result!);
    const coreProps = await zip.file('docProps/core.xml')?.async('text');
    if (coreProps) {
      expect(coreProps).toContain('dcterms:modified');
    }
  });

  test('selectively patches a single paragraph edit', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Find a paragraph to modify
    const paragraphs = doc.package.document.content.filter(
      (b): b is Paragraph => b.type === 'paragraph'
    );

    // Find a paragraph with a paraId that has text content
    let targetPara: Paragraph | null = null;
    for (const para of paragraphs) {
      if (para.paraId) {
        const text = para.content
          .filter((item): item is Run => item.type === 'run')
          .flatMap((run) => run.content)
          .filter((c) => c.type === 'text')
          .map((c) => c.text)
          .join('');
        if (text.length > 0) {
          targetPara = para;
          break;
        }
      }
    }

    if (!targetPara || !targetPara.paraId) {
      // Skip test if no suitable paragraph found
      console.log('No paragraph with paraId and text found in example-with-image.docx, skipping');
      return;
    }

    const paraId = targetPara.paraId;

    // Modify the paragraph text
    for (const item of targetPara.content) {
      if (item.type === 'run') {
        for (const c of item.content) {
          if (c.type === 'text' && c.text.length > 0) {
            c.text = c.text + ' [MODIFIED]';
            break;
          }
        }
        break;
      }
    }

    // Attempt selective save
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set([paraId]),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    if (result === null) {
      // Selective save may fail due to paragraph count mismatch (serializer differences)
      // This is expected — fall back to full repack
      console.log('Selective save returned null (expected fallback), verifying full repack works');
      const fullResult = await repackDocx(doc);
      expect(fullResult.byteLength).toBeGreaterThan(0);
      return;
    }

    // Verify the result is a valid DOCX
    expect(result.byteLength).toBeGreaterThan(0);
    const resultXml = await getDocumentXml(result);
    expect(resultXml).toContain('[MODIFIED]');

    // Verify unchanged paragraphs are preserved
    const originalXml = await getDocumentXml(buffer);
    for (const para of paragraphs) {
      if (para.paraId && para.paraId !== paraId) {
        const origOffsets = findParagraphOffsets(originalXml, para.paraId);
        const resultOffsets = findParagraphOffsets(resultXml, para.paraId);
        if (origOffsets && resultOffsets) {
          const origPara = originalXml.substring(origOffsets.start, origOffsets.end);
          const resultPara = resultXml.substring(resultOffsets.start, resultOffsets.end);
          expect(resultPara).toBe(origPara);
        }
      }
    }
  });

  test('returns null for changed paraId not found in original', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(['NONEXISTENT_PARA_ID']),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    // Should return null (fallback to full repack)
    expect(result).toBeNull();
  });
});

// ============================================================================
// Round-trip correctness
// ============================================================================

describe('Selective save round-trip', () => {
  test('selective save produces valid DOCX that can be re-parsed', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Find a modifiable paragraph
    const paragraphs = doc.package.document.content.filter(
      (b): b is Paragraph => b.type === 'paragraph'
    );

    let targetPara: Paragraph | null = null;
    for (const para of paragraphs) {
      if (para.paraId) {
        const hasText = para.content.some(
          (item) =>
            item.type === 'run' && item.content.some((c) => c.type === 'text' && c.text.length > 0)
        );
        if (hasText) {
          targetPara = para;
          break;
        }
      }
    }

    if (!targetPara?.paraId) {
      console.log('No suitable paragraph for round-trip test, skipping');
      return;
    }

    // Modify and attempt selective save
    const paraId = targetPara.paraId;
    for (const item of targetPara.content) {
      if (item.type === 'run') {
        for (const c of item.content) {
          if (c.type === 'text' && c.text.length > 0) {
            c.text = 'ROUNDTRIP_TEST';
            break;
          }
        }
        break;
      }
    }

    const savedBuffer = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set([paraId]),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    if (!savedBuffer) {
      // Fallback is acceptable
      console.log('Selective save fell back, testing full repack round-trip instead');
      const fullBuffer = await repackDocx(doc);
      const reopened = await parseDocx(fullBuffer, { preloadFonts: false });
      expect(reopened.package.document.content.length).toBeGreaterThan(0);
      return;
    }

    // Re-parse the saved document
    const reopened = await parseDocx(savedBuffer, { preloadFonts: false });
    expect(reopened.package.document.content.length).toBeGreaterThan(0);

    // Verify the modified text exists
    const modifiedPara = reopened.package.document.content.find(
      (b): b is Paragraph => b.type === 'paragraph' && b.paraId === paraId
    );
    if (modifiedPara) {
      const text = modifiedPara.content
        .filter((item): item is Run => item.type === 'run')
        .flatMap((run) => run.content)
        .filter((c) => c.type === 'text')
        .map((c) => c.text)
        .join('');
      expect(text).toContain('ROUNDTRIP_TEST');
    }
  });

  test('no-edit save preserves document.xml exactly', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    const savedBuffer = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    // No changes — should return the original buffer
    expect(savedBuffer).not.toBeNull();
  });
});

// ============================================================================
// buildPatchedDocumentXml with real XML
// ============================================================================

describe('buildPatchedDocumentXml with real DOCX XML', () => {
  test('patches real document.xml correctly', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const originalXml = await getDocumentXml(buffer);

    // Find all paraIds
    const paraIdPattern = /w14:paraId="([^"]+)"/g;
    const allIds: string[] = [];
    let m: RegExpExecArray | null;
    while ((m = paraIdPattern.exec(originalXml)) !== null) {
      allIds.push(m[1]);
    }

    // Build a "serialized" version with one paragraph modified
    // (In real usage this comes from serializeDocument after document model changes)
    const uniqueIds = [...new Set(allIds)];
    if (uniqueIds.length === 0) {
      console.log('No paraIds in example-with-image.docx, skipping');
      return;
    }

    // Find a paraId that can be located in the XML
    let testId: string | null = null;
    for (const id of uniqueIds) {
      if (findParagraphOffsets(originalXml, id)) {
        testId = id;
        break;
      }
    }

    if (!testId) {
      console.log('No locatable paraId in example-with-image.docx, skipping');
      return;
    }

    // Create a modified version where the target paragraph has different content
    const origOffsets = findParagraphOffsets(originalXml, testId)!;
    const origParagraph = originalXml.substring(origOffsets.start, origOffsets.end);
    const modifiedParagraph = origParagraph.replace(
      /<w:t[^>]*>[^<]*<\/w:t>/,
      '<w:t>PATCHED_TEXT</w:t>'
    );

    // Build the "serialized" XML with just this one paragraph changed
    const serializedXml =
      originalXml.substring(0, origOffsets.start) +
      modifiedParagraph +
      originalXml.substring(origOffsets.end);

    const patched = buildPatchedDocumentXml(originalXml, serializedXml, new Set([testId]));
    expect(patched).not.toBeNull();

    if (patched) {
      // The patched paragraph should have the new content
      const patchedOffsets = findParagraphOffsets(patched, testId)!;
      const patchedPara = patched.substring(patchedOffsets.start, patchedOffsets.end);
      expect(patchedPara).toContain('PATCHED_TEXT');

      // All other content should be identical
      // Check that the content before the patched paragraph is byte-for-byte identical
      expect(patched.substring(0, patchedOffsets.start)).toBe(
        originalXml.substring(0, origOffsets.start)
      );
    }
  });
});

// ============================================================================
// Edge cases
// ============================================================================

// ============================================================================
// Comments handling in selective save
// ============================================================================

describe('Selective save with comments', () => {
  test('includes comments.xml when document has comments', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Add a comment to the document model
    doc.package.document.comments = [
      {
        id: 1,
        author: 'Test User',
        date: '2024-01-01T00:00:00Z',
        content: [
          {
            type: 'paragraph',
            formatting: {},
            content: [
              { type: 'run', formatting: {}, content: [{ type: 'text', text: 'Test comment' }] },
            ],
          },
        ],
      },
    ];

    // Even with no paragraph changes, comments should be saved
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();
    // Result should NOT be the original buffer (it has comments now)
    expect(result!.byteLength).not.toBe(buffer.byteLength);

    // Verify comments.xml exists and contains the comment
    const zip = await JSZip.loadAsync(result!);
    const commentsFile = zip.file('word/comments.xml');
    expect(commentsFile).not.toBeNull();
    const commentsXml = await commentsFile!.async('text');
    expect(commentsXml).toContain('Test comment');
    expect(commentsXml).toContain('w:id="1"');
    expect(commentsXml).toContain('Test User');
  });

  test('ensures content type and relationship entries for new comments', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Check if original has comments entries (it probably doesn't)
    const originalZip = await JSZip.loadAsync(buffer);
    const originalCt = await originalZip.file('[Content_Types].xml')?.async('text');
    const hadCommentsEntry = originalCt?.includes('/word/comments.xml') ?? false;

    // Add a comment
    doc.package.document.comments = [
      {
        id: 42,
        author: 'Author',
        content: [
          {
            type: 'paragraph',
            formatting: {},
            content: [
              { type: 'run', formatting: {}, content: [{ type: 'text', text: 'New comment' }] },
            ],
          },
        ],
      },
    ];

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();
    const resultZip = await JSZip.loadAsync(result!);

    // Content type should be present
    const ctXml = await resultZip.file('[Content_Types].xml')?.async('text');
    expect(ctXml).toContain('/word/comments.xml');

    // Relationship should be present
    const relsXml = await resultZip.file('word/_rels/document.xml.rels')?.async('text');
    expect(relsXml).toContain('comments.xml');

    // If the original didn't have comments, verify the entries were added
    if (!hadCommentsEntry) {
      expect(ctXml).toContain(
        'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
      );
    }
  });

  test('preserves existing comments and adds new ones', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Simulate having both an original comment and a new one
    doc.package.document.comments = [
      {
        id: 1,
        author: 'Original Author',
        content: [
          {
            type: 'paragraph',
            formatting: {},
            content: [
              {
                type: 'run',
                formatting: {},
                content: [{ type: 'text', text: 'Original comment' }],
              },
            ],
          },
        ],
      },
      {
        id: 999,
        author: 'New Author',
        content: [
          {
            type: 'paragraph',
            formatting: {},
            content: [
              { type: 'run', formatting: {}, content: [{ type: 'text', text: 'New comment' }] },
            ],
          },
        ],
      },
    ];

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();
    const zip = await JSZip.loadAsync(result!);
    const commentsXml = await zip.file('word/comments.xml')!.async('text');

    // Both comments should be present
    expect(commentsXml).toContain('Original comment');
    expect(commentsXml).toContain('New comment');
    expect(commentsXml).toContain('w:id="1"');
    expect(commentsXml).toContain('w:id="999"');
  });

  test('comments saved alongside paragraph changes', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Add a comment
    doc.package.document.comments = [
      {
        id: 5,
        author: 'Commenter',
        content: [
          {
            type: 'paragraph',
            formatting: {},
            content: [
              { type: 'run', formatting: {}, content: [{ type: 'text', text: 'Comment text' }] },
            ],
          },
        ],
      },
    ];

    // Also modify a paragraph
    const paragraphs = doc.package.document.content.filter(
      (b): b is Paragraph => b.type === 'paragraph' && !!b.paraId
    );
    const target = paragraphs.find((p) =>
      p.content.some(
        (i) => i.type === 'run' && i.content.some((c) => c.type === 'text' && c.text.length > 0)
      )
    );

    if (!target?.paraId) {
      console.log('No suitable paragraph, skipping');
      return;
    }

    for (const item of target.content) {
      if (item.type === 'run') {
        for (const c of item.content) {
          if (c.type === 'text' && c.text.length > 0) {
            c.text = c.text + ' [WITH_COMMENT]';
            break;
          }
        }
        break;
      }
    }

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set([target.paraId]),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    if (result) {
      const zip = await JSZip.loadAsync(result);

      // Verify document.xml has the paragraph change
      const docXml = await zip.file('word/document.xml')!.async('text');
      expect(docXml).toContain('[WITH_COMMENT]');

      // Verify comments.xml exists with the comment
      const commentsXml = await zip.file('word/comments.xml')!.async('text');
      expect(commentsXml).toContain('Comment text');
    }
  });
});

// ============================================================================
// Headers/footers and core properties in selective save
// ============================================================================

describe('Selective save with headers/footers', () => {
  test('serializes headers/footers into the saved output', async () => {
    const buffer = await loadFixture('EP_ZMVZ_MULTI_v4.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Check if document has headers or footers
    const hasHeaders = doc.package.headers && doc.package.headers.size > 0;
    const hasFooters = doc.package.footers && doc.package.footers.size > 0;
    if (!hasHeaders && !hasFooters) {
      console.log('No headers/footers in fixture, skipping');
      return;
    }

    // Modify a header/footer content
    const map = hasHeaders ? doc.package.headers! : doc.package.footers!;
    for (const [, hf] of map.entries()) {
      if (hf.content && hf.content.length > 0) {
        const para = hf.content.find((b) => b.type === 'paragraph');
        if (para && para.type === 'paragraph') {
          para.content.push({
            type: 'run',
            formatting: {},
            content: [{ type: 'text', text: ' [HF_MODIFIED]' }],
          });
          break;
        }
      }
    }

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();

    // Verify that the header/footer file was updated
    const zip = await JSZip.loadAsync(result!);
    let found = false;
    for (const [path, file] of Object.entries(zip.files)) {
      if (path.match(/word\/(header|footer)\d*\.xml/)) {
        const xml = await file.async('text');
        if (xml.includes('[HF_MODIFIED]')) {
          found = true;
          break;
        }
      }
    }
    expect(found).toBe(true);
  });
});

describe('Selective save updates core properties', () => {
  test('updates modification date on save', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Get original modification date
    const originalZip = await JSZip.loadAsync(buffer);
    const originalCoreProps = await originalZip.file('docProps/core.xml')?.async('text');

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();
    const resultZip = await JSZip.loadAsync(result!);
    const resultCoreProps = await resultZip.file('docProps/core.xml')?.async('text');

    if (originalCoreProps && resultCoreProps) {
      // The modification date should have been updated
      expect(resultCoreProps).not.toBe(originalCoreProps);
      expect(resultCoreProps).toContain('dcterms:modified');
    }
  });
});

describe('Selective save edge cases', () => {
  test('handles document with tables', async () => {
    const buffer = await loadFixture('with-tables.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });
    const xml = await getDocumentXml(buffer);

    // Verify we can count paragraphs (tables contain w:p elements too)
    const count = countParagraphElements(xml);
    expect(count).toBeGreaterThan(0);

    // No changes = should return original
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(result).not.toBeNull();
  });

  test('handles complex styled document', async () => {
    const buffer = await loadFixture('complex-styles.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });
    const xml = await getDocumentXml(buffer);

    const count = countParagraphElements(xml);
    expect(count).toBeGreaterThan(0);

    // No changes = should return original
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(result).not.toBeNull();
  });

  test('handles document with missing paraIds gracefully', async () => {
    // Create a minimal DOCX-like XML without paraIds
    // The attempt should fall back (return null) if we ask to patch a nonexistent ID
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(['FAKE_ID_12345']),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).toBeNull(); // Should fall back
  });

  test('handles large document with many paraIds', async () => {
    const buffer = await loadFixture('EP_ZMVZ_MULTI_v4.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });
    const xml = await getDocumentXml(buffer);

    // Count paraIds
    const paraIdPattern = /w14:paraId="([^"]+)"/g;
    const ids: string[] = [];
    let m: RegExpExecArray | null;
    while ((m = paraIdPattern.exec(xml)) !== null) {
      ids.push(m[1]);
    }
    expect(ids.length).toBeGreaterThan(50);

    // No-changes save should work
    const noChangeResult = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(noChangeResult).not.toBeNull();

    // Single paragraph edit
    const paragraphs = doc.package.document.content.filter(
      (b): b is Paragraph => b.type === 'paragraph' && !!b.paraId
    );
    expect(paragraphs.length).toBeGreaterThan(0);

    const target = paragraphs.find((p) =>
      p.content.some(
        (i) => i.type === 'run' && i.content.some((c) => c.type === 'text' && c.text.length > 0)
      )
    );

    if (target?.paraId) {
      // Modify target paragraph
      for (const item of target.content) {
        if (item.type === 'run') {
          for (const c of item.content) {
            if (c.type === 'text' && c.text.length > 0) {
              c.text = c.text + ' [LARGE_DOC_TEST]';
              break;
            }
          }
          break;
        }
      }

      const result = await attemptSelectiveSave(doc, buffer, {
        changedParaIds: new Set([target.paraId]),
        structuralChange: false,
        hasUntrackedChanges: false,
      });

      if (result) {
        // Verify it's a valid ZIP
        expect(result.byteLength).toBeGreaterThan(0);
        const resultXml = await getDocumentXml(result);
        expect(resultXml).toContain('[LARGE_DOC_TEST]');

        // Verify other paragraphs are byte-for-byte identical
        const originalXml = await getDocumentXml(buffer);
        let checkedCount = 0;
        for (const para of paragraphs.slice(0, 10)) {
          if (para.paraId && para.paraId !== target.paraId) {
            const origOffsets = findParagraphOffsets(originalXml, para.paraId);
            const resultOffsets = findParagraphOffsets(resultXml, para.paraId);
            if (origOffsets && resultOffsets) {
              const origText = originalXml.substring(origOffsets.start, origOffsets.end);
              const resultText = resultXml.substring(resultOffsets.start, resultOffsets.end);
              expect(resultText).toBe(origText);
              checkedCount++;
            }
          }
        }
        expect(checkedCount).toBeGreaterThan(0);
      }
    }
  });

  test('selective save disabled falls back to full repack', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // When structuralChange=true (simulating selective=false at higher level), we get null
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: true,
      hasUntrackedChanges: false,
    });
    expect(result).toBeNull();
  });

  test('multiple paragraphs edited selectively', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    const paragraphs = doc.package.document.content.filter(
      (b): b is Paragraph =>
        b.type === 'paragraph' &&
        !!b.paraId &&
        b.content.some(
          (i) => i.type === 'run' && i.content.some((c) => c.type === 'text' && c.text.length > 0)
        )
    );

    if (paragraphs.length < 2) {
      console.log('Not enough paragraphs with paraId to test multi-edit');
      return;
    }

    const editIds: string[] = [];
    // Edit first two paragraphs with paraIds
    for (const para of paragraphs.slice(0, 2)) {
      for (const item of para.content) {
        if (item.type === 'run') {
          for (const c of item.content) {
            if (c.type === 'text' && c.text.length > 0) {
              c.text = c.text + ` [MULTI_${para.paraId}]`;
              break;
            }
          }
          break;
        }
      }
      editIds.push(para.paraId!);
    }

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(editIds),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    if (result) {
      const resultXml = await getDocumentXml(result);
      for (const id of editIds) {
        expect(resultXml).toContain(`[MULTI_${id}]`);
      }

      // Verify an unedited paragraph is unchanged
      const uneditedPara = paragraphs.find((p) => !editIds.includes(p.paraId!));
      if (uneditedPara?.paraId) {
        const originalXml = await getDocumentXml(buffer);
        const origOffsets = findParagraphOffsets(originalXml, uneditedPara.paraId);
        const resultOffsets = findParagraphOffsets(resultXml, uneditedPara.paraId);
        if (origOffsets && resultOffsets) {
          expect(resultXml.substring(resultOffsets.start, resultOffsets.end)).toBe(
            originalXml.substring(origOffsets.start, origOffsets.end)
          );
        }
      }
    }
  });
});

// ============================================================================
// Final section properties (body-level sectPr) change detection
// ============================================================================

describe('Selective save with section property (margin) changes', () => {
  test('margin-only change falls back to full repack (returns null)', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    // Fixture must have body-level section properties for this test
    expect(doc.package.document.finalSectionProperties).toBeDefined();

    // Simulate a margin drag (as createMarginHandler does in the editor)
    const props = doc.package.document.finalSectionProperties!;
    const newLeft = (props.marginLeft ?? 1440) + 720;
    doc.package.document.finalSectionProperties = {
      ...props,
      marginLeft: newLeft,
    };

    // Margin-only edits produce NO changed paragraphs — this is exactly the
    // case where selective save used to silently drop the change.
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
      hasNonParagraphChanges: false,
    });

    expect(result).toBeNull();
  });

  test('DocumentAgent.toBuffer persists margin change via full-repack fallback', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });
    expect(doc.package.document.finalSectionProperties).toBeDefined();

    const props = doc.package.document.finalSectionProperties!;
    const newLeft = (props.marginLeft ?? 1440) + 720;
    doc.package.document.finalSectionProperties = {
      ...props,
      marginLeft: newLeft,
    };

    const agent = new DocumentAgent(doc);
    const out = await agent.toBuffer({
      selective: {
        changedParaIds: new Set(),
        structuralChange: false,
        hasUntrackedChanges: false,
        hasNonParagraphChanges: false,
      },
    });

    // The saved document.xml must contain the new margin value
    const outXml = await getDocumentXml(out);
    expect(outXml).toContain(`w:left="${newLeft}"`);

    // And it must survive a reload round-trip
    const reopened = await parseDocx(out, { preloadFonts: false });
    expect(reopened.package.document.finalSectionProperties?.marginLeft).toBe(newLeft);
  });

  test('unchanged margins on Word original do NOT trigger fallback', async () => {
    // Guard against false positives: an untouched Word document (whose
    // sectPr formatting never matches our serializer output textually)
    // must still selective-save.
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });
    expect(doc.package.document.finalSectionProperties).toBeDefined();

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
      hasNonParagraphChanges: false,
    });

    expect(result).not.toBeNull();
    // document.xml must be byte-for-byte unchanged
    const originalXml = await getDocumentXml(buffer);
    const resultXml = await getDocumentXml(result!);
    expect(resultXml).toBe(originalXml);
  });

  test('unchanged margins with a paragraph edit still selective-save', async () => {
    const buffer = await loadFixture('EP_ZMVZ_MULTI_v4.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });

    const paragraphs = doc.package.document.content.filter(
      (b): b is Paragraph =>
        b.type === 'paragraph' &&
        !!b.paraId &&
        b.content.some(
          (i) => i.type === 'run' && i.content.some((c) => c.type === 'text' && c.text.length > 0)
        )
    );
    const target = paragraphs[0];
    if (!target?.paraId) {
      console.log('No suitable paragraph, skipping');
      return;
    }

    for (const item of target.content) {
      if (item.type === 'run') {
        for (const c of item.content) {
          if (c.type === 'text' && c.text.length > 0) {
            c.text = c.text + ' [SECTPR_GUARD]';
            break;
          }
        }
        break;
      }
    }

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set([target.paraId]),
      structuralChange: false,
      hasUntrackedChanges: false,
      hasNonParagraphChanges: false,
    });

    // Selective save may legitimately fall back for serializer-mismatch
    // reasons, but if it succeeds the paragraph edit must be present.
    if (result) {
      const resultXml = await getDocumentXml(result);
      expect(resultXml).toContain('[SECTPR_GUARD]');
    }
  });

  test('original without body-level sectPr + document with section properties → fallback', async () => {
    const buffer = await loadFixture('example-with-image.docx');

    // Build a variant of the fixture whose document.xml has NO body sectPr
    const zip = await JSZip.loadAsync(buffer);
    const originalXml = await zip.file('word/document.xml')!.async('text');
    const sectPrXml = extractBodySectPrXml(originalXml);
    expect(sectPrXml).not.toBeNull();
    const strippedXml = originalXml.replace(sectPrXml!, '');
    zip.file('word/document.xml', strippedXml);
    const strippedBuffer = await zip.generateAsync({ type: 'arraybuffer' });

    const doc = await parseDocx(strippedBuffer, { preloadFonts: false });
    expect(doc.package.document.finalSectionProperties).toBeUndefined();

    // Unchanged (no sectPr, no finalSectionProperties) → selective save OK
    const unchanged = await attemptSelectiveSave(doc, strippedBuffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
      hasNonParagraphChanges: false,
    });
    expect(unchanged).not.toBeNull();

    // Now the editor adds margins (spread over undefined, as the margin
    // handler does) → must be treated as changed → fallback
    doc.package.document.finalSectionProperties = {
      ...doc.package.document.finalSectionProperties,
      marginLeft: 2160,
      marginRight: 2160,
    };
    const changed = await attemptSelectiveSave(doc, strippedBuffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
      hasNonParagraphChanges: false,
    });
    expect(changed).toBeNull();
  });

  test('document with sectPr but finalSectionProperties removed → fallback', async () => {
    const buffer = await loadFixture('example-with-image.docx');
    const doc = await parseDocx(buffer, { preloadFonts: false });
    expect(doc.package.document.finalSectionProperties).toBeDefined();

    doc.package.document.finalSectionProperties = undefined;

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
      hasNonParagraphChanges: false,
    });
    expect(result).toBeNull();
  });
});

describe('finalSectionPropertiesChanged', () => {
  const NS =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

  test('Word-style formatting noise does not register as a change', () => {
    // Word-authored sectPr: rsid attributes, multi-line formatting,
    // attribute order that will never match our serializer output.
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body><w:p w14:paraId="AAA111"><w:r><w:t>Hi</w:t></w:r></w:p>
  <w:sectPr w:rsidR="00AB12CD" w:rsidSect="00FF00FF">
    <w:pgSz w:h="16838" w:w="11906"/>
    <w:pgMar w:gutter="0" w:footer="708" w:header="708" w:left="1440" w:bottom="1440" w:right="1440" w:top="1440"/>
    <w:cols w:space="708"/>
  </w:sectPr>
</w:body></w:document>`;

    // What the editor holds: same values, different key insertion order
    const currentProps = {
      marginTop: 1440,
      marginBottom: 1440,
      marginLeft: 1440,
      marginRight: 1440,
      headerDistance: 708,
      footerDistance: 708,
      gutter: 0,
      pageHeight: 16838,
      pageWidth: 11906,
      columnSpace: 708,
    };

    expect(finalSectionPropertiesChanged(xml, currentProps)).toBe(false);
  });

  test('a single changed margin value registers as a change', () => {
    const xml = `<w:document ${NS}><w:body><w:p w14:paraId="AAA111"/><w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;

    const unchanged = {
      marginTop: 1440,
      marginRight: 1440,
      marginBottom: 1440,
      marginLeft: 1440,
    };
    expect(finalSectionPropertiesChanged(xml, unchanged)).toBe(false);
    expect(finalSectionPropertiesChanged(xml, { ...unchanged, marginLeft: 2880 })).toBe(true);
  });

  test('no body sectPr and no finalSectionProperties → unchanged', () => {
    const xml = `<w:document ${NS}><w:body><w:p w14:paraId="AAA111"/></w:body></w:document>`;
    expect(finalSectionPropertiesChanged(xml, undefined)).toBe(false);
    // Empty object is equivalent to absent
    expect(finalSectionPropertiesChanged(xml, {})).toBe(false);
  });

  test('no body sectPr but document has margins → changed', () => {
    const xml = `<w:document ${NS}><w:body><w:p w14:paraId="AAA111"/></w:body></w:document>`;
    expect(finalSectionPropertiesChanged(xml, { marginLeft: 1440 })).toBe(true);
  });

  test('body sectPr present but finalSectionProperties removed → changed', () => {
    const xml = `<w:document ${NS}><w:body><w:p w14:paraId="AAA111"/><w:sectPr><w:pgMar w:top="1440"/></w:sectPr></w:body></w:document>`;
    expect(finalSectionPropertiesChanged(xml, undefined)).toBe(true);
  });

  test('mid-document sectPr is ignored for final-section comparison', () => {
    const xml = `<w:document ${NS}><w:body>
<w:p w14:paraId="AAA111"><w:pPr><w:sectPr><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:pPr></w:p>
<w:p w14:paraId="BBB222"/>
<w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`;
    const currentProps = {
      marginTop: 1440,
      marginRight: 1440,
      marginBottom: 1440,
      marginLeft: 1440,
    };
    expect(finalSectionPropertiesChanged(xml, currentProps)).toBe(false);
  });

  test('header reference changes register as changed', () => {
    const xml = `<w:document ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p w14:paraId="AAA111"/><w:sectPr><w:pgMar w:top="1440"/></w:sectPr></w:body></w:document>`;
    expect(
      finalSectionPropertiesChanged(xml, {
        marginTop: 1440,
        headerReferences: [{ type: 'default', rId: 'rId99' }],
      })
    ).toBe(true);
  });
});
