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
 * A cell 400 wide holding a line, an empty paragraph, a floating table of the cell's width anchored `offset` above
 * its paragraph, and that paragraph: the probe tabelle-flottanti F1 of the bench, as the Mosaico of Golden Group has it
 * on its page 317.
 */
function floatingTableInCell(offset: number) {
    Reflect.set(FontCache, '_context', null);
    FontCache.invalidateMetrics(() => true);
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
        font: '',
        measureText(this: CanvasRenderingContext2D, value: string) {
            return { width: value.length * 4, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 };
        },
    } as unknown as CanvasRenderingContext2D);
    const cell = (content: string) => `${T.TABLE_CELL_START}${content}${T.SECTION_BREAK}${T.TABLE_CELL_END}`;
    const inner = `${T.TABLE_START}${Array.from({ length: 3 }, (_, row) => `${T.TABLE_ROW_START}${cell(`R${row}${P}`)}${T.TABLE_ROW_END}`).join('')}${T.TABLE_END}`;
    const content = `Prima${P}${P}${inner}Dopo${P}`;
    const outer = `${T.TABLE_START}${T.TABLE_ROW_START}${cell(content)}${T.TABLE_ROW_END}${T.TABLE_END}`;
    const dataStream = `${outer}${P}${T.SECTION_BREAK}`;
    const empty = outer.indexOf(`${P}${P}`) + 1;
    const after = outer.indexOf(`Dopo${P}`) + 4;
    const margin = { start: { v: 0 }, end: { v: 0 }, top: { v: 0 }, bottom: { v: 0 } };
    const table = (tableId: string, rows: number, wrap: boolean): ITable => ({
        tableId,
        align: TableAlignmentType.START,
        indent: { v: 0 },
        textWrap: wrap ? TableTextWrapType.WRAP : TableTextWrapType.NONE,
        position: wrap
            ? { positionH: { relativeFrom: ObjectRelativeFromH.MARGIN, posOffset: 0 }, positionV: { relativeFrom: ObjectRelativeFromV.PARAGRAPH, posOffset: offset } }
            : { positionH: { relativeFrom: ObjectRelativeFromH.PAGE }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE } },
        dist: { distT: 0, distB: 0, distL: wrap ? 9 : 0, distR: wrap ? 9 : 0 },
        size: { type: TableSizeType.SPECIFIED, width: { v: wrap ? 396 : 400 } },
        layout: TableLayoutType.FIXED,
        cellMargin: margin,
        tableColumns: [{ size: { type: TableSizeType.SPECIFIED, width: { v: wrap ? 396 : 400 } } }],
        tableRows: Array.from({ length: rows }, () => ({ cantSplit: BooleanNumber.FALSE, trHeight: { hRule: TableRowHeightRule.AUTO, val: { v: 0 } }, tableCells: [{ margin }] })),
    } as unknown as ITable);
    const bed = createParagraphLayoutTestBed(outer, {
        body: {
            dataStream,
            paragraphs: Array.from(dataStream).flatMap((token, startIndex) => token === P ? [{ startIndex, paragraphId: `p-${startIndex}`, paragraphStyle: { lineSpacing: 16, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE } }] : []),
            sectionBreaks: Array.from(dataStream).flatMap((token, startIndex) => token === T.SECTION_BREAK ? [{ startIndex, sectionId: `s-${startIndex}` }] : []),
            tables: [{ tableId: 'outer', startIndex: 0, endIndex: outer.length }, { tableId: 'inner', startIndex: outer.indexOf(inner), endIndex: outer.indexOf(inner) + inner.length }],
        },
        tableSource: { outer: table('outer', 1, false), inner: table('inner', 3, true) },
        documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, textStyle: { ff: 'Arial', fs: 11 }, pageSize: { width: 440, height: 600 }, marginLeft: 20, marginRight: 20, marginTop: 20, marginBottom: 20 },
    });
    return { bed, empty, after };
}

function lineTop(page: IDocumentSkeletonPage, paragraphIndex: number): number | undefined {
    for (const section of page.sections) {
        for (const column of section.columns) {
            for (const line of column.lines) {
                if (line.paragraphIndex === paragraphIndex) {
                    return section.top + line.top;
                }
            }
        }
    }
    return undefined;
}

describe('a floating table with no room beside it', () => {
    afterEach(() => vi.restoreAllMocks());

    // Word places the table from where its paragraph stands, 8 above it here, and lays out below the table the lines
    // before it that the table covers: the empty paragraph before it goes below it, and its paragraph after that.
    it('lays out below itself the line before it that it covers', () => {
        const { bed, empty, after } = floatingTableInCell(-8);
        const skeleton = DocumentSkeleton.create(bed.viewModel, bed.ctx.docsConfig.localeService);
        try {
            skeleton.calculate();
            const page = skeleton.getSkeletonData()!.pages[0];
            const cellPage = [...page.skeTables.values()][0].rows[0].cells[0];
            const inner = cellPage.skeTables.get('inner')!;
            // the table starts 8 above where the empty paragraph ended, its paragraph's place
            expect(inner.top).toBeCloseTo(16 + 16 - 8, 1);
            expect(lineTop(cellPage, empty)).toBeGreaterThanOrEqual(inner.top + inner.height - 0.01);
            expect(lineTop(cellPage, after)).toBeCloseTo(lineTop(cellPage, empty)! + 16, 1);
        } finally {
            bed.viewModel.dispose();
            bed.dataModel.dispose();
        }
    });
});
