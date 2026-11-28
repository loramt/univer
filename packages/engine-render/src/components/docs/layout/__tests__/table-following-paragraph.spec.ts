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
import type { IDocumentSkeletonPage } from '../../../../basics/i-document-skeleton-cached';
import { createParagraphLayoutTestBed } from '../block/paragraph/__tests__/create-paragraph-layout-test-bed';
import { DocumentSkeleton } from '../doc-skeleton';
import { FontCache } from '../shaping-engine/font-cache';

const T = DataStreamTreeTokenType;
const P = T.PARAGRAPH;

/**
 * `filler` lines of 16, then a table of `rows` rows of 16 and the empty paragraph after it, `markHeight` high, then a
 * last paragraph. The body is 404 high.
 */
function tableThenEmptyParagraph(filler: number, rows: number, markHeight: number) {
    Reflect.set(FontCache, '_context', null);
    FontCache.invalidateMetrics(() => true);
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
        font: '',
        measureText(this: CanvasRenderingContext2D, value: string) {
            return { width: value.length * 4, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 };
        },
    } as unknown as CanvasRenderingContext2D);
    const cell = (content: string) => `${T.TABLE_CELL_START}${content}${T.SECTION_BREAK}${T.TABLE_CELL_END}`;
    const table = `${T.TABLE_START}${Array.from({ length: rows }, (_, row) => `${T.TABLE_ROW_START}${cell(`R${row}${P}`)}${T.TABLE_ROW_END}`).join('')}${T.TABLE_END}`;
    const fill = Array.from({ length: filler }, (_, index) => `F${index}${P}`).join('');
    const content = `${fill}${table}${P}Dopo${P}`;
    const dataStream = `${content}${T.SECTION_BREAK}`;
    const mark = fill.length + table.length;
    const margin = { start: { v: 4 }, end: { v: 4 }, top: { v: 0 }, bottom: { v: 0 } };
    const bed = createParagraphLayoutTestBed(content.slice(0, -1), {
        body: {
            dataStream,
            paragraphs: Array.from(dataStream).flatMap((token, startIndex) => token === P
                ? [{ startIndex, paragraphId: `p-${startIndex}`, paragraphStyle: { lineSpacing: startIndex === mark ? markHeight : 16, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE } }]
                : []),
            sectionBreaks: Array.from(dataStream).flatMap((token, startIndex) => token === T.SECTION_BREAK ? [{ startIndex, sectionId: `s-${startIndex}` }] : []),
            tables: [{ tableId: 'table', startIndex: fill.length, endIndex: fill.length + table.length }],
        },
        tableSource: {
            table: {
                tableId: 'table',
                align: TableAlignmentType.START,
                indent: { v: 0 },
                textWrap: TableTextWrapType.NONE,
                position: { positionH: { relativeFrom: ObjectRelativeFromH.PAGE }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE } },
                dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
                size: { type: TableSizeType.SPECIFIED, width: { v: 400 } },
                layout: TableLayoutType.FIXED,
                cellMargin: margin,
                tableColumns: [{ size: { type: TableSizeType.SPECIFIED, width: { v: 400 } } }],
                tableRows: Array.from({ length: rows }, () => ({
                    cantSplit: BooleanNumber.FALSE,
                    trHeight: { hRule: TableRowHeightRule.AUTO, val: { v: 0 } },
                    tableCells: [{ margin }],
                })),
            } as unknown as ITable,
        },
        documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, textStyle: { ff: 'Arial', fs: 11 }, pageSize: { width: 600, height: 444 }, marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20 },
    });
    return { bed, mark, after: mark + 1 };
}

function pageOf(pages: IDocumentSkeletonPage[], index: number): number {
    return pages.findIndex((page) => page.sections.some((section) => section.columns.some((column) => column.lines.some((line) =>
        line.divides.some((divide) => divide.glyphGroup.some((glyph) => glyph.count > 0 && glyph.raw === P && line.paragraphIndex === index))))));
}

describe('the paragraph after a table that ends at the foot of a page', () => {
    afterEach(() => vi.restoreAllMocks());

    // The first page holds the filler and 24 rows (400 of 404); the second the other rows, and below them 36 are left.
    // A mark 40 high does not fit there and goes to the third page, also when the table opened the second page: Word
    // moves it there (probe vuoto-dopo-tabella-piede, the empty paragraph after the table of the Mosaico's page 188).
    // A mark 24 high stays.
    it.each([
        [1, 47, 40, 2],
        [1, 47, 24, 1],
        [3, 20, 40, 1],
        [3, 20, 24, 0],
    ])('with %i filler lines and %i rows a mark %i high is on page %i', (filler, rows, markHeight, expected) => {
        const { bed, mark } = tableThenEmptyParagraph(filler, rows, markHeight);
        const skeleton = DocumentSkeleton.create(bed.viewModel, bed.ctx.docsConfig.localeService);
        try {
            skeleton.calculate();
            const pages = skeleton.getSkeletonData()!.pages;
            expect(pageOf(pages, mark)).toBe(expected);
        } finally {
            bed.viewModel.dispose();
            bed.dataModel.dispose();
        }
    });
});
