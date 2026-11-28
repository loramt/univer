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

import type { ITable } from '@univerjs/core';
import {
    BooleanNumber,
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
import { createParagraphLayoutTestBed } from '../block/paragraph/__tests__/create-paragraph-layout-test-bed';
import { lineRoom } from '../block/paragraph/line-room';
import { DocumentSkeleton } from '../doc-skeleton';
import { takeLayoutWarnings } from '../layout-warnings';
import { FontCache } from '../shaping-engine/font-cache';
import { clearFontCreateConfigCache } from '../tools';

/** Three paragraphs of one line each, exactly lineHeight high, in a body 60 high. */
function pagesFor(lineHeight: number): number {
    Reflect.set(FontCache, '_context', null);
    FontCache.invalidateMetrics(() => true);
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
        font: '',
        measureText(this: CanvasRenderingContext2D, value: string) {
            return { width: value.length * 4, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 };
        },
    } as unknown as CanvasRenderingContext2D);
    const content = 'A\rB\rC';
    const dataStream = `${content}\r\n`;
    const bed = createParagraphLayoutTestBed(content, {
        body: {
            dataStream,
            paragraphs: [...dataStream.matchAll(/\r/g)].map((match) => ({
                startIndex: match.index!,
                paragraphId: `p-${match.index}`,
                paragraphStyle: { lineSpacing: lineHeight, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE },
            })),
            sectionBreaks: [{ sectionId: 'main', startIndex: dataStream.length - 1 }],
        },
        documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, pageSize: { width: 200, height: 100 }, marginTop: 20, marginBottom: 20, marginLeft: 20, marginRight: 20 },
    });
    const skeleton = DocumentSkeleton.create(bed.viewModel, bed.ctx.docsConfig.localeService);
    try {
        skeleton.calculate();
        return skeleton.getSkeletonData()!.pages.length;
    } finally {
        skeleton.dispose();
        bed.viewModel.dispose();
        bed.dataModel.dispose();
    }
}

describe('lines at the foot of a page of a Word document', () => {
    afterEach(() => vi.restoreAllMocks());

    // Word compares where a line ends with its room exactly: the three lines end at the end of the body, 60 high, and
    // stay; a tenth of a twip past it (15 twips to the pixel), the third goes to the next page.
    it.each([[20, 1], [20 - 0.1 / 45, 1], [20 + 0.1 / 45, 2]])('with lines %s high the three lines take %i page(s)', (lineHeight, pages) => {
        expect(pagesFor(lineHeight)).toBe(pages);
    });
});

/**
 * A table of one row opening a body 60 high, its cell holding a paragraph of four lines `lineHeight` high, then a line
 * after the table. Returns how many lines of the row are on the first page, and the number of pages.
 */
function cellParagraph(lineHeight: number): { lines: number; pages: number } {
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
        font: '',
        measureText(this: CanvasRenderingContext2D, value: string) {
            // 4 px a character at 11 pt, in proportion at the size the font cache measures with
            const match = this.font.match(/([\d.]+)(pt|px)/);
            const ratio = (match ? Number(match[1]) * (match[2] === 'px' ? 0.75 : 1) : 11) / 11;
            return { width: value.length * 4 * ratio, fontBoundingBoxAscent: 8 * ratio, fontBoundingBoxDescent: 2 * ratio, actualBoundingBoxAscent: 8 * ratio, actualBoundingBoxDescent: 2 * ratio };
        },
    } as unknown as CanvasRenderingContext2D);
    Reflect.set(FontCache, '_context', null);
    FontCache.invalidateMetrics(() => true);
    clearFontCreateConfigCache();
    const T = DataStreamTreeTokenType;
    // four lines of words 6 and 9 characters long, 4 px each, in a cell 160 wide: the last line holds three words
    const text = Array.from({ length: 9 }, (_, index) => `Riga${index} ${'x'.repeat(8)}`).join(' ');
    const table = `${T.TABLE_START}${T.TABLE_ROW_START}${T.TABLE_CELL_START}${text}${T.PARAGRAPH}${T.SECTION_BREAK}${T.TABLE_CELL_END}${T.TABLE_ROW_END}${T.TABLE_END}`;
    const content = `${table}Dopo${T.PARAGRAPH}`;
    const dataStream = `${content}${T.SECTION_BREAK}`;
    const cellParagraphIndex = table.indexOf(T.PARAGRAPH);
    const margin = { start: { v: 0 }, end: { v: 0 }, top: { v: 0 }, bottom: { v: 0 } };
    const bed = createParagraphLayoutTestBed(content, {
        body: {
            dataStream,
            paragraphs: Array.from(dataStream).flatMap((token, startIndex) => token === T.PARAGRAPH
                ? [{ startIndex, paragraphId: `p-${startIndex}`, paragraphStyle: { lineSpacing: startIndex === cellParagraphIndex ? lineHeight : 14, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE } }]
                : []),
            sectionBreaks: Array.from(dataStream).flatMap((token, startIndex) => token === T.SECTION_BREAK ? [{ startIndex, sectionId: `s-${startIndex}` }] : []),
            tables: [{ tableId: 'table', startIndex: 0, endIndex: table.length }],
        },
        tableSource: {
            table: {
                tableId: 'table',
                align: TableAlignmentType.START,
                indent: { v: 0 },
                textWrap: TableTextWrapType.NONE,
                position: { positionH: { relativeFrom: ObjectRelativeFromH.PAGE }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE } },
                dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
                size: { type: TableSizeType.SPECIFIED, width: { v: 160 } },
                layout: TableLayoutType.FIXED,
                cellMargin: margin,
                tableColumns: [{ size: { type: TableSizeType.SPECIFIED, width: { v: 160 } } }],
                tableRows: [{ cantSplit: BooleanNumber.FALSE, trHeight: { hRule: TableRowHeightRule.AUTO, val: { v: 0 } }, tableCells: [{ margin }] }],
            } as unknown as ITable,
        },
        documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, textStyle: { ff: 'Arial', fs: 11 }, pageSize: { width: 200, height: 100 }, marginTop: 20, marginBottom: 20, marginLeft: 20, marginRight: 20 },
    });
    const skeleton = DocumentSkeleton.create(bed.viewModel, bed.ctx.docsConfig.localeService);
    try {
        skeleton.calculate();
        const pages = skeleton.getSkeletonData()!.pages;
        const cell = [...pages[0].skeTables.values()][0]?.rows[0]?.cells[0];
        return { lines: cell?.sections[0].columns[0].lines.length ?? 0, pages: pages.length };
    } finally {
        skeleton.dispose();
        bed.viewModel.dispose();
        bed.dataModel.dispose();
    }
}

describe('the last line of a paragraph in a part of a table row', () => {
    afterEach(() => vi.restoreAllMocks());

    // Four lines of 15 end at the end of the body, 60 high: the row keeps them, and the line after the table goes to the
    // second page. Lines of 15.005 end 0.3 twip past it: the last one goes to the second page, as in Word (the Mosaico of
    // Golden Group: the line of a cell 0.24 twip past its room), and widow control takes the third along. The checks of
    // the room of the line agree: where one let such a line stay and another did not, the line was laid out again
    // until the stack ran out.
    it.each([[15, 4], [15.005, 2]])('with lines %s high the row keeps %i lines on the first page', (lineHeight, kept) => {
        const { lines, pages } = cellParagraph(lineHeight);
        expect(lines).toBe(kept);
        expect(pages).toBe(2);
    });

    // Two checks that disagree on purpose: the one before laying the end of the paragraph out again always finds it
    // past its room, the one of where the line goes lets it stay. The line is laid out again inside its own layout
    // until the bound, then stays as it is, and the layout ends, with a warning each time the row is laid out.
    it('ends with a warning when two checks of the room of a line disagree', () => {
        vi.spyOn(lineRoom, 'cellParagraphEndOverflows').mockReturnValue(true);
        const { lines, pages } = cellParagraph(15);
        const warnings = takeLayoutWarnings();
        expect([lines, pages]).toEqual([4, 2]);
        expect(warnings.length).toBeGreaterThan(0);
        for (const warning of warnings) {
            expect(warning).toContain('was laid out again 16 times inside its own layout');
        }
    });
});
