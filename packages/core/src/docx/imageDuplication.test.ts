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
