/**
 * Regression tests for indent persistence round-trips.
 *
 * Scenario: a Word-authored DOCX has paragraphs with negative indents
 * (`<w:ind w:left="-566" w:right="-607"/>`). The user drags the ruler's
 * indent markers back to the margin (indent = 0). This must persist:
 *
 * 1. setIndentLeft(0) must store an explicit 0 (previously clamped to null,
 *    so the original -566 round-tripped back via _originalFormatting).
 * 2. fromProseDoc must not drop 0 values via falsy `||` filtering.
 * 3. The serializer must emit w:left="0" (an explicit override), not omit it.
 */

import { describe, test, expect } from 'bun:test';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Transaction } from 'prosemirror-state';
import { toProseDoc } from './toProseDoc';
import { fromProseDoc } from './fromProseDoc';
import { schema } from '../schema';
import { ParagraphExtension } from '../extensions/core/ParagraphExtension';
import { serializeParagraph } from '../../docx/serializer/paragraphSerializer';
import type { Document, Paragraph } from '../../types/document';
import type { ParagraphFormatting } from '../../types/formatting';

const runtime = ParagraphExtension().onSchemaReady({ schema });
const commands = runtime.commands!;

function makeDocument(formatting: ParagraphFormatting): Document {
  const paragraph: Paragraph = {
    type: 'paragraph',
    paraId: '11111111',
    formatting,
    content: [{ type: 'run', content: [{ type: 'text', text: 'Hello world' }] }],
  };
  return {
    package: {
      document: { content: [paragraph] },
    },
  } as Document;
}

/** Create an editor state with the cursor inside the first paragraph. */
function stateFor(doc: Document): EditorState {
  const pmDoc = toProseDoc(doc);
  let state = EditorState.create({ doc: pmDoc });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
  return state;
}

function apply(state: EditorState, command: ReturnType<typeof commands.setIndentLeft>) {
  let next = state;
  const result = command(state, (tr: Transaction) => {
    next = state.apply(tr);
  });
  expect(result).toBe(true);
  return next;
}

function firstParagraph(doc: Document): Paragraph {
  const block = doc.package.document.content[0];
  expect(block.type).toBe('paragraph');
  return block as Paragraph;
}

describe('indent round-trip — negative indents from Word', () => {
  const NEGATIVE_IND: ParagraphFormatting = { indentLeft: -566, indentRight: -607 };

  test('negative indents reach PM attrs on load', () => {
    const state = stateFor(makeDocument(NEGATIVE_IND));
    const para = state.doc.child(0);
    expect(para.attrs.indentLeft).toBe(-566);
    expect(para.attrs.indentRight).toBe(-607);
  });

  test('setIndentLeft(0) stores explicit 0 and serializes as w:left="0"', () => {
    let state = stateFor(makeDocument(NEGATIVE_IND));
    state = apply(state, commands.setIndentLeft(0));

    const para = state.doc.child(0);
    expect(para.attrs.indentLeft).toBe(0);
    // Write-through to the lossless serialization base
    expect((para.attrs._originalFormatting as ParagraphFormatting).indentLeft).toBe(0);

    const out = fromProseDoc(state.doc);
    const formatting = firstParagraph(out).formatting!;
    expect(formatting.indentLeft).toBe(0);
    // Untouched right indent is preserved
    expect(formatting.indentRight).toBe(-607);

    const xml = serializeParagraph(firstParagraph(out));
    expect(xml).toContain('w:left="0"');
    expect(xml).toContain('w:right="-607"');
    expect(xml).not.toContain('-566');
  });

  test('setIndentRight(0) clears the negative right indent on save', () => {
    let state = stateFor(makeDocument(NEGATIVE_IND));
    state = apply(state, commands.setIndentRight(0));

    const out = fromProseDoc(state.doc);
    const xml = serializeParagraph(firstParagraph(out));
    expect(xml).toContain('w:right="0"');
    expect(xml).toContain('w:left="-566"');
    expect(xml).not.toContain('-607');
  });

  test('setIndentLeft with a positive value persists via _originalFormatting', () => {
    let state = stateFor(makeDocument(NEGATIVE_IND));
    state = apply(state, commands.setIndentLeft(720));

    const out = fromProseDoc(state.doc);
    const xml = serializeParagraph(firstParagraph(out));
    expect(xml).toContain('w:left="720"');
    expect(xml).not.toContain('-566');
  });

  test('setIndentFirstLine(0, false) removes an original hanging indent', () => {
    let state = stateFor(makeDocument({ indentFirstLine: 283, hangingIndent: true }));
    state = apply(state, commands.setIndentFirstLine(0, false));

    const out = fromProseDoc(state.doc);
    const xml = serializeParagraph(firstParagraph(out));
    expect(xml).not.toContain('w:hanging');
    expect(xml).toContain('w:firstLine="0"');
  });

  test('decreaseIndent from a negative indent lands on explicit 0', () => {
    let state = stateFor(makeDocument(NEGATIVE_IND));
    state = apply(state, commands.decreaseIndent());

    const para = state.doc.child(0);
    expect(para.attrs.indentLeft).toBe(0);

    const out = fromProseDoc(state.doc);
    const xml = serializeParagraph(firstParagraph(out));
    expect(xml).toContain('w:left="0"');
    expect(xml).not.toContain('-566');
  });

  test('untouched paragraph round-trips its original negative indents losslessly', () => {
    const state = stateFor(makeDocument(NEGATIVE_IND));
    const out = fromProseDoc(state.doc);
    const xml = serializeParagraph(firstParagraph(out));
    expect(xml).toContain('w:left="-566"');
    expect(xml).toContain('w:right="-607"');
  });
});

describe('fromProseDoc fallback — 0 values survive (|| regression)', () => {
  test('editor-created paragraph with indentLeft 0 keeps explicit 0', () => {
    const pmDoc = schema.node('doc', null, [
      schema.node('paragraph', { indentLeft: 0 }, [schema.text('x')]),
    ]);
    const out = fromProseDoc(pmDoc);
    const formatting = firstParagraph(out).formatting;
    expect(formatting).toBeDefined();
    expect(formatting!.indentLeft).toBe(0);
  });

  test('editor-created paragraph with spaceBefore/spaceAfter 0 keeps explicit 0', () => {
    const pmDoc = schema.node('doc', null, [
      schema.node('paragraph', { spaceBefore: 0, spaceAfter: 0 }, [schema.text('x')]),
    ]);
    const out = fromProseDoc(pmDoc);
    const formatting = firstParagraph(out).formatting;
    expect(formatting).toBeDefined();
    expect(formatting!.spaceBefore).toBe(0);
    expect(formatting!.spaceAfter).toBe(0);
  });
});
