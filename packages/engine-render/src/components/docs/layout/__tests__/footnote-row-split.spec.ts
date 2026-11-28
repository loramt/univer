/**
 * Copyright 2023-present DreamNum Co., Ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { describe, expect, it } from 'vitest';
import type { IDocumentSkeletonGlyph } from '../../../../basics/i-document-skeleton-cached';
import type { INoteReferenceLine } from '../footnote-layout';
import { findNoteRowCapacity } from '../footnote-layout';

// Lines of a row with one note reference each, 14 high, in a cell with 3 of bottom margin.
function referenceLines(tops: number[]): INoteReferenceLine[] {
    return tops.map((top, index) => ({
        top,
        bottom: top + 14 + 3,
        split: top + 3,
        atTop: top === 0,
        glyphs: [{ noteId: `note-${index}` } as IDocumentSkeletonGlyph],
    }));
}

// A page whose body ends at 480 without notes; each note takes 40 below the text, and a note that has no room
// below the text is postponed to the next page.
function pageWithNotes(noteHeight = 40, body = 480, postponed: string[] = []) {
    return (kept: readonly INoteReferenceLine[], bottom: number) => {
        const limit = body - kept.length * noteHeight;
        const ids = kept.flatMap((line) => line.glyphs.map((glyph) => glyph.noteId!));
        return { limit, deferredIds: new Set(limit < bottom ? ids : ids.filter((id) => postponed.includes(id))) };
    };
}

describe('findNoteRowCapacity', () => {
    it('keeps the whole row when its notes fit below it', () => {
        const lines = referenceLines([10, 60]);
        // two notes leave the body down to 400, the row ends at 120
        expect(findNoteRowCapacity(lines, 120, pageWithNotes())).toBe(400);
    });

    it('ends the row before the first reference line whose note has no room, as Word does', () => {
        // Word's Mosaico of Golden Group, page 215: six references; with all of them on the page their notes would
        // leave no room, with the first four the notes fit, and Word splits the row before the fifth reference.
        const lines = referenceLines([10, 130, 250, 262, 310, 440]);
        // four notes: the body may reach 320, the fourth reference line ends at 279, the fifth starts at 310
        expect(findNoteRowCapacity(lines, 457, pageWithNotes())).toBe(313);
    });

    it('keeps the lines before the split together with the bottom margin of their cell', () => {
        // two text lines (0-23.5) before the only reference line, whose note does not fit: the fragment keeps both
        // lines, 26.5 high with the cell's margins, not 23.5
        const lines = referenceLines([26.5 - 3]);
        expect(findNoteRowCapacity(lines, 40, pageWithNotes(40, 50))).toBeCloseTo(26.5);
    });

    it('sends on the reference lines of notes that a continuing note postpones', () => {
        const lines = referenceLines([10, 60, 110]);
        // a note continued from the previous page fills the page's notes: the second and third notes cannot start
        expect(findNoteRowCapacity(lines, 127, pageWithNotes(40, 480, ['note-1', 'note-2']))).toBe(63);
    });

    it('keeps a reference line that opens the fragment, whatever happens to its note', () => {
        // its note does not fit below it (it will continue on the next page): the line stays on the page
        const lines = referenceLines([0, 30]);
        expect(findNoteRowCapacity(lines, 47, pageWithNotes(40, 50))).toBe(17);
    });

    it('ends the part before the first reference line of a row that can move, when its note has no room', () => {
        // the row starts below other content of the page: Word moves it to the next page (the Mosaico of Golden Group,
        // page 183), so its first reference line does not stay whatever happens to its note
        const lines = referenceLines([0, 30]);
        expect(findNoteRowCapacity(lines, 47, pageWithNotes(40, 50), true)).toBe(3);
    });

    it('keeps the reference line of a continuing note in a row that can move, also with the limit rounded', () => {
        // the note continues on the next page, so the body ends with its reference line: the limit is the line's bottom,
        // measured back from the page (the verifier's probe v8r-lunga-40: 16.00999999999995 for 16.01). The line stays
        // and the part ends with it; without the tolerance the part held nothing and the row went to the next page.
        const lines = referenceLines([0]).map((line) => ({ ...line, bottom: 16.01 }));
        const capacity = findNoteRowCapacity(lines, 16.01, () => ({ limit: 16.00999999999995, deferredIds: new Set<string>() }), true);
        expect(capacity).toBeCloseTo(16.01, 6);
    });
});
