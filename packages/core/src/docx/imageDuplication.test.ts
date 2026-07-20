/**
 * Regression test for image duplication bug.
 *
 * When a document containing embedded images is parsed and re-saved without
 * any changes, images should not be duplicated in the output ZIP.
 *
 * Root cause: collectNewImages() treated all data-URL src values as "new"
 * images, including images that were loaded from the original file (which
 * always have both src=data: AND a pre-existing rId). On each save, these
 * were re-added as new media files while the originals were also preserved.
 */

import { describe, test, expect } from 'bun:test';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { parseDocx } from './parser';
import { repackDocx } from './rezip';

const FIXTURES_DIR = path.resolve(__dirname, '../../../../e2e/fixtures');

function loadFixture(name: string): ArrayBuffer {
  const buf = fs.readFileSync(path.join(FIXTURES_DIR, name));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

async function countMediaFiles(buffer: ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer);
  return Object.keys(zip.files).filter((p) => p.startsWith('word/media/') && !zip.files[p].dir);
}

async function countImageRelationships(buffer: ArrayBuffer, relsPath: string): Promise<number> {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(relsPath);
  if (!file) return 0;
  const xml = await file.async('text');
  const matches = xml.match(/Type="[^"]*\/image"/g);
  return matches ? matches.length : 0;
}

describe('image duplication on save', () => {
test('re-saving a document without changes should NOT duplicate images', async () => {
    const original = loadFixture('example-with-image.docx');

    const originalMedia = await countMediaFiles(original);
    const originalRelCount = await countImageRelationships(original, 'word/_rels/document.xml.rels');
    console.log('Original media files:', originalMedia);
    console.log('Original image relationships:', originalRelCount);

    const doc = await parseDocx(original);
    const saved = await repackDocx(doc);

    const savedMedia = await countMediaFiles(saved);
    const savedRelCount = await countImageRelationships(saved, 'word/_rels/document.xml.rels');
    console.log('Saved media files:', savedMedia);
    console.log('Saved image relationships:', savedRelCount);

    // After fix: media count and relationship count must be identical
    expect(savedMedia.length).toBe(originalMedia.length);
    expect(savedRelCount).toBe(originalRelCount);
  });

  test('re-saving three times does not grow the document', async () => {
    const original = loadFixture('example-with-image.docx');
    const originalMedia = await countMediaFiles(original);

    let doc = await parseDocx(original);
    for (let i = 0; i < 3; i++) {
      const saved = await repackDocx(doc);
      doc = await parseDocx(saved);
    }

    const finalSaved = await repackDocx(doc);
    const finalMedia = await countMediaFiles(finalSaved);
    console.log(`After 3 round-trips: ${originalMedia.length} → ${finalMedia.length} media files`);

    expect(finalMedia.length).toBe(originalMedia.length);
  });
});

// 1x1 red pixel PNG
const TINY_PNG_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function insertImageParagraph(doc: Awaited<ReturnType<typeof parseDocx>>, rId: string): void {
  doc.package.document.content.push({
    type: 'paragraph',
    formatting: {},
    content: [
      {
        type: 'run',
        formatting: {},
        content: [
          {
            type: 'drawing',
            image: {
              type: 'image',
              rId,
              src: TINY_PNG_DATA_URL,
              size: { width: 9525, height: 9525 },
              wrap: { type: 'inline' },
            },
          },
        ],
      },
    ],
  });
}

async function parsedImageCount(buffer: ArrayBuffer): Promise<number> {
  const reparsed = await parseDocx(buffer);
  let count = 0;
  for (const block of reparsed.package.document.content) {
    if (block.type !== 'paragraph') continue;
    for (const item of block.content) {
      if (item.type !== 'run') continue;
      for (const c of item.content) {
        if (c.type === 'drawing' && c.image?.src) count++;
      }
    }
  }
  return count;
}

describe('inserted image persistence across save', () => {
  test('image with empty rId is embedded and survives reparse', async () => {
    const doc = await parseDocx(loadFixture('empty.docx'));
    insertImageParagraph(doc, '');

    const saved = await repackDocx(doc);

    const media = await countMediaFiles(saved);
    expect(media.length).toBe(1);
    expect(await countImageRelationships(saved, 'word/_rels/document.xml.rels')).toBe(1);
    expect(await parsedImageCount(saved)).toBe(1);
  });

  test('image with fake/unregistered rId is still embedded (regression: image lost on reopen)', async () => {
    const doc = await parseDocx(loadFixture('empty.docx'));
    // Mimics the placeholder rIds historically stamped at insert time
    // (`rId_img_<timestamp>`): truthy, but with no relationship or media
    // behind it. Before the fix this serialized a dangling r:embed and the
    // image vanished on reopen.
    insertImageParagraph(doc, `rId_img_${1234567890}`);

    const saved = await repackDocx(doc);

    const media = await countMediaFiles(saved);
    expect(media.length).toBe(1);
    expect(await parsedImageCount(saved)).toBe(1);
  });

  test('re-saving when the editor re-presents the image as new does not grow the package', async () => {
    const doc = await parseDocx(loadFixture('empty.docx'));
    insertImageParagraph(doc, '');

    const firstSave = await repackDocx(doc);
    expect((await countMediaFiles(firstSave)).length).toBe(1);

    // Simulate the React flow: document content is rebuilt from ProseMirror
    // on every save, so the rId assigned during the first repack is lost and
    // the same data: URL comes back marked as new.
    doc.originalBuffer = firstSave;
    const para = doc.package.document.content[doc.package.document.content.length - 1];
    if (para.type === 'paragraph' && para.content[0]?.type === 'run') {
      const drawing = para.content[0].content[0];
      if (drawing.type === 'drawing') drawing.image.rId = '';
    }

    const secondSave = await repackDocx(doc);

    // Byte-dedupe must reuse the existing media file instead of adding a copy
    expect((await countMediaFiles(secondSave)).length).toBe(1);
    expect(await parsedImageCount(secondSave)).toBe(1);
  });
});
