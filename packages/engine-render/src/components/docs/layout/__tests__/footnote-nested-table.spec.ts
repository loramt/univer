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

import type { IDocumentNote, ITable } from '@univerjs/core';
import {
    BooleanNumber,
    CustomRangeType,
    DataStreamTreeTokenType,
    DocumentFlavor,
    ObjectRelativeFromH,
    ObjectRelativeFromV,
    SpacingRule,
    TableAlignmentType,
    TableLayoutType,
    TableRowHeightRule,
    TableSizeType,
    TableTextWrapType,
} from '@univerjs/core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { IDocumentSkeletonPage } from '../../../../basics/i-document-skeleton-cached';
import { createParagraphLayoutTestBed } from '../block/paragraph/__tests__/create-paragraph-layout-test-bed';
import { DocumentSkeleton } from '../doc-skeleton';
import { FontCache } from '../shaping-engine/font-cache';

const T = DataStreamTreeTokenType;
const P = T.PARAGRAPH;
const NOTE_LINES = [3, 7, 4, 2, 9];

/**
 * The verifier's B1 (probes v7-blocco and v7-annidate-pagine-c55): a fixed table with one row, whose second cell has
 * `filler` paragraphs and a table of 60 rows nested in it (the first row cannot split), with references to five notes
 * in its rows 2, 5, 8, 40 and 45.
 */
function nestedTableWithNotes(filler: number, pageHeight: number) {
    Reflect.set(FontCache, '_context', null);
    FontCache.invalidateMetrics(() => true);
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
        font: '',
        measureText(this: CanvasRenderingContext2D, value: string) {
            return { width: value.length * 4, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 };
        },
    } as unknown as CanvasRenderingContext2D);
    const references = new Map([[2, 0], [5, 1], [8, 2], [40, 3], [45, 4]]);
    const cell = (content: string) => `${T.TABLE_CELL_START}${content}${T.SECTION_BREAK}${T.TABLE_CELL_END}`;
    const rows = Array.from({ length: 60 }, (_, row) => `${T.TABLE_ROW_START}${cell(`T${row}${P}`)}${cell(`${row === 0 ? 'lungo '.repeat(40) : `misura ${row}`}${references.has(row) ? '￼' : ''}${P}`)}${T.TABLE_ROW_END}`).join('');
    const inner = `${T.TABLE_START}${rows}${T.TABLE_END}`;
    const fill = Array.from({ length: filler }, (_, index) => `F${index}${P}`).join('');
    const outer = `${T.TABLE_START}${T.TABLE_ROW_START}${cell(`Etichetta${P}`)}${cell(`${fill}${inner}${P}Z8${P}`)}${T.TABLE_ROW_END}${T.TABLE_END}`;
    const dataStream = `${outer}${P}${T.SECTION_BREAK}`;
    const margin = { start: { v: 4 }, end: { v: 4 }, top: { v: 0 }, bottom: { v: 0 } };
    const border = { s: 1, w: 1, cl: { rgb: '#000000' } };
    const table = (tableId: string, grid: number[], rowCount: number, cantSplitFirst: boolean) => ({
        tableId,
        align: TableAlignmentType.START,
        indent: { v: 0 },
        textWrap: TableTextWrapType.NONE,
        position: { positionH: { relativeFrom: ObjectRelativeFromH.PAGE }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE } },
        dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
        size: { type: TableSizeType.SPECIFIED, width: { v: grid.reduce((sum, width) => sum + width, 0) } },
        layout: TableLayoutType.FIXED,
        cellMargin: margin,
        tableColumns: grid.map((width) => ({ size: { type: TableSizeType.SPECIFIED, width: { v: width } } })),
        tableRows: Array.from({ length: rowCount }, (_, row) => ({
            cantSplit: cantSplitFirst && row === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            trHeight: { hRule: TableRowHeightRule.AUTO, val: { v: 0 } },
            tableCells: grid.map(() => ({ margin, borderTop: border, borderBottom: border, borderLeft: border, borderRight: border })),
        })),
    } as unknown as ITable);
    const noteStream = (lines: number) => `${Array.from({ length: lines }, (_, i) => `Nota riga ${i + 1}${P}`).join('')}\n`;
    const marks = [...dataStream.matchAll(/￼/g)].map((match, index) => ({ index: match.index!, id: `note-${index}` }));
    const bed = createParagraphLayoutTestBed(outer, {
        body: {
            dataStream,
            paragraphs: Array.from(dataStream).flatMap((token, startIndex) => token === P ? [{ startIndex, paragraphId: `p-${startIndex}`, paragraphStyle: { lineSpacing: 16, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE } }] : []),
            sectionBreaks: Array.from(dataStream).flatMap((token, startIndex) => token === T.SECTION_BREAK ? [{ startIndex, sectionId: `s-${startIndex}` }] : []),
            tables: [{ tableId: 'outer', startIndex: 0, endIndex: outer.length }, { tableId: 'inner', startIndex: outer.indexOf(inner), endIndex: outer.indexOf(inner) + inner.length }],
            customRanges: marks.map(({ index, id }) => ({ rangeId: `ref-${id}`, rangeType: CustomRangeType.FOOTNOTE, startIndex: index, endIndex: index, wholeEntity: true, properties: { noteId: id } })),
        },
        tableSource: { outer: table('outer', [100, 400], 1, false), inner: table('inner', [80, 300], 60, true) },
        documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, textStyle: { ff: 'Arial', fs: 11 }, pageSize: { width: 600, height: pageHeight }, marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20 },
        notes: Object.fromEntries(marks.map(({ id }, index) => [id, { type: 'footnote' as const, noteId: id, body: {
            dataStream: noteStream(NOTE_LINES[index]),
            paragraphs: [...noteStream(NOTE_LINES[index]).matchAll(/\r/g)].map((match, k) => ({ paragraphId: `${id}-${k}`, startIndex: match.index! })),
            sectionBreaks: [{ sectionId: `${id}-section`, startIndex: noteStream(NOTE_LINES[index]).length - 1 }],
        } } as IDocumentNote])),
    });
    return bed;
}

function noteText(pages: IDocumentSkeletonPage[], noteId: string): string {
    return pages.flatMap((page) => page.notes ?? []).filter((fragment) => fragment.noteId === noteId).flatMap((fragment) => fragment.page.sections.flatMap((section) => section.columns.flatMap((column) =>
        column.lines.flatMap((line) => line.divides.flatMap((divide) => divide.glyphGroup))))).filter((glyph) => glyph.count > 0).map((glyph) => glyph.raw).join('');
}

describe('notes of a table nested in a row that splits', () => {
    afterEach(() => vi.restoreAllMocks());

    // Without the correction these layouts never end: a part of the outer row, measured with its notes, comes out
    // taller than its measure, its notes cannot start on its page, and the part is moved to the next page, where it
    // opens the page just the same ("Maximum call stack size exceeded"). The part stays, and the notes that cannot
    // start on its page start on the next one.
    it.each([[400, 0], [400, 21], [500, 1], [500, 17], [500, 27]])('ends with every note laid out (page %i, %i paragraphs before)', (pageHeight, filler) => {
        const bed = nestedTableWithNotes(filler, pageHeight);
        const skeleton = DocumentSkeleton.create(bed.viewModel, bed.ctx.docsConfig.localeService);
        try {
            expect(() => skeleton.calculate()).not.toThrow();
            const pages = skeleton.getSkeletonData()!.pages;
            expect(pages.length).toBeLessThan(10);
            NOTE_LINES.forEach((lines, index) => {
                expect(noteText(pages, `note-${index}`)).toBe(`${Array.from({ length: lines }, (_, i) => `Nota riga ${i + 1}\r`).join('')}\n`);
            });
        } finally {
            bed.viewModel.dispose();
            bed.dataModel.dispose();
        }
    });
});
