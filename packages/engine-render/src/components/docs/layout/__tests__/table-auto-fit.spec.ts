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
    NumberUnitType,
    ObjectRelativeFromH,
    ObjectRelativeFromV,
    SpacingRule,
    TableAlignmentType,
    TableLayoutType,
    TableRowHeightRule,
    TableSizeType,
    TableTextWrapType,
} from '@univerjs/core';
import { describe, expect, it, vi } from 'vitest';
import type { IWordAutoFit } from '../table-auto-fit';
import { createParagraphLayoutTestBed } from '../block/paragraph/__tests__/create-paragraph-layout-test-bed';
import { resolveTableLayoutColumns } from '../block/table';
import { DocumentSkeleton } from '../doc-skeleton';
import { FontCache } from '../shaping-engine/font-cache';
import { resolveTableAutoFitColumns } from '../table-auto-fit';
import { clearFontCreateConfigCache } from '../tools';

describe('autofit tables of Word', () => {
    /** Lays out a one-row table of these texts (4 px a character, no padding) on a page 600 px wide. */
    function columnsOf(texts: string[], autoFit: IWordAutoFit, grid = texts.map(() => 200)) {
        const previousContext = Reflect.get(FontCache, '_context');
        Reflect.set(FontCache, '_context', null);
        FontCache.invalidateMetrics(() => true);
        const context = {
            font: '',
            measureText(this: CanvasRenderingContext2D, value: string) {
                // 4 px a character at 11 pt, in proportion at the size the font cache measures with
                const match = this.font.match(/([\d.]+)(pt|px)/);
                const ratio = (match ? Number(match[1]) * (match[2] === 'px' ? 0.75 : 1) : 11) / 11;
                return { width: value.length * 4 * ratio, fontBoundingBoxAscent: 8 * ratio, fontBoundingBoxDescent: 2 * ratio, actualBoundingBoxAscent: 8 * ratio, actualBoundingBoxDescent: 2 * ratio };
            },
        } as unknown as CanvasRenderingContext2D;
        const canvas = vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(context);
        const T = DataStreamTreeTokenType;
        const row = `${T.TABLE_ROW_START}${texts.map((value) => `${T.TABLE_CELL_START}${value}${T.PARAGRAPH}${T.SECTION_BREAK}${T.TABLE_CELL_END}`).join('')}${T.TABLE_ROW_END}`;
        const stream = `${T.TABLE_START}${row}${T.TABLE_END}`;
        const dataStream = `${stream}${T.PARAGRAPH}${T.SECTION_BREAK}`;
        const zero = { start: { v: 0 }, end: { v: 0 }, top: { v: 0 }, bottom: { v: 0 } };
        // as the converter of Word documents writes it: fixed, with its own columns, and the preferred widths
        const source = {
            tableId: 'word-fit',
            align: TableAlignmentType.START,
            indent: { v: 0 },
            textWrap: TableTextWrapType.NONE,
            position: { positionH: { relativeFrom: ObjectRelativeFromH.PAGE }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE } },
            dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
            size: { type: TableSizeType.SPECIFIED, width: { v: grid.reduce((sum, width) => sum + width, 0) } },
            layout: TableLayoutType.FIXED,
            cellMargin: zero,
            tableColumns: grid.map((width) => ({ size: { type: TableSizeType.SPECIFIED, width: { v: width } } })),
            tableRows: [{ trHeight: { hRule: TableRowHeightRule.AUTO, val: { v: 0 } }, tableCells: texts.map(() => ({ margin: zero })) }],
            wordAutoFit: autoFit,
        } as ITable;
        const bed = createParagraphLayoutTestBed(stream, {
            body: {
                paragraphs: Array.from(dataStream).flatMap((token, startIndex) => token === T.PARAGRAPH ? [{ startIndex, paragraphId: `p-${startIndex}`, paragraphStyle: { lineSpacing: 16, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE } }] : []),
                sectionBreaks: Array.from(dataStream).flatMap((token, startIndex) => token === T.SECTION_BREAK ? [{ startIndex, sectionId: `s-${startIndex}` }] : []),
                tables: [{ tableId: source.tableId, startIndex: 0, endIndex: stream.length }],
            },
            tableSource: { [source.tableId]: source },
            documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, textStyle: { ff: 'Arial', fs: 11 }, pageSize: { width: 600, height: 800 }, marginLeft: 0, marginRight: 0 },
        });
        const skeleton = DocumentSkeleton.create(bed.viewModel, bed.sectionBreakConfig.localeService);
        try {
            skeleton.calculate();
            const table = skeleton.getSkeletonData()!.pages[0].skeTables.get(source.tableId)!;
            return table.tableSource.tableColumns.map((column) => Math.round(column.size.width.v * 100) / 100);
        } finally {
            skeleton.dispose();
            bed.viewModel.dispose();
            bed.dataModel.dispose();
            canvas.mockRestore();
            Reflect.set(FontCache, '_context', previousContext);
            FontCache.invalidateMetrics(() => true);
            clearFontCreateConfigCache();
        }
    }
    const auto = { type: 'auto' as const, value: 0 };
    const long = 'aa bb '.repeat(33).trim();

    it('gives up the room of the columns without a width first, by what they have above their narrowest', () => {
        // a table of the whole width: a long text (788 px on one line), a short one (20 px) and a word (12 px); 220 px
        // too many, taken by what each has above its narrowest word (780, 12, 0)
        expect(columnsOf([long, 'xx yy', 'zzz'], { width: { type: 'pct', value: 1 }, columns: [auto, auto, auto] }))
            .toEqual([571.33, 16.67, 12]);
    });

    it('gives the columns in percent their share while the shares add up to 100%', () => {
        // three columns at 50%: the third gets none and keeps the room of its word
        const half = { type: 'pct' as const, value: 0.5 };
        expect(columnsOf(['aaa', 'bbb', 'parole'], { width: { type: 'pct', value: 1 }, columns: [half, half, half] }))
            .toEqual([288, 288, 24]);
    });

    it('gives the room left to the columns without a width, by their widest', () => {
        expect(columnsOf(['aa', 'bbbb', 'cccccc'], { width: { type: 'pct', value: 1 }, columns: [auto, auto, auto] }))
            .toEqual([100, 200, 300]);
    });

    /** An autofit table of these texts nested in the second cell of a fixed table, after what the cell has before it. */
    function nestedTableOf(texts: string[], before = '') {
        const previousContext = Reflect.get(FontCache, '_context');
        Reflect.set(FontCache, '_context', null);
        FontCache.invalidateMetrics(() => true);
        const context = {
            font: '',
            measureText(this: CanvasRenderingContext2D, value: string) {
                return { width: value.length * 4, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 };
            },
        } as unknown as CanvasRenderingContext2D;
        const canvas = vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(context);
        const T = DataStreamTreeTokenType;
        const cell = (content: string) => `${T.TABLE_CELL_START}${content}${T.PARAGRAPH}${T.SECTION_BREAK}${T.TABLE_CELL_END}`;
        const inner = `${T.TABLE_START}${T.TABLE_ROW_START}${texts.map(cell).join('')}${T.TABLE_ROW_END}${T.TABLE_END}`;
        const outer = `${T.TABLE_START}${T.TABLE_ROW_START}${cell('esterna')}${cell(before + inner)}${T.TABLE_ROW_END}${T.TABLE_END}`;
        const dataStream = `${outer}${T.PARAGRAPH}${T.SECTION_BREAK}`;
        const zero = { start: { v: 0 }, end: { v: 0 }, top: { v: 0 }, bottom: { v: 0 } };
        const table = (tableId: string, grid: number[], autoFit?: IWordAutoFit) => ({
            tableId,
            align: TableAlignmentType.START,
            indent: { v: 0 },
            textWrap: TableTextWrapType.NONE,
            position: { positionH: { relativeFrom: ObjectRelativeFromH.PAGE }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE } },
            dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
            size: { type: TableSizeType.SPECIFIED, width: { v: grid.reduce((sum, width) => sum + width, 0) } },
            layout: TableLayoutType.FIXED,
            cellMargin: zero,
            tableColumns: grid.map((width) => ({ size: { type: TableSizeType.SPECIFIED, width: { v: width } } })),
            tableRows: [{ trHeight: { hRule: TableRowHeightRule.AUTO, val: { v: 0 } }, tableCells: grid.map(() => ({ margin: zero })) }],
            ...(autoFit ? { wordAutoFit: autoFit } : {}),
        } as ITable);
        const nestedAutoFit = { width: { type: 'auto' as const, value: 0 }, columns: texts.map(() => auto) };
        const bed = createParagraphLayoutTestBed(outer, {
            body: {
                paragraphs: Array.from(dataStream).flatMap((token, startIndex) => token === T.PARAGRAPH ? [{ startIndex, paragraphId: `p-${startIndex}`, paragraphStyle: { lineSpacing: 16, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE } }] : []),
                sectionBreaks: Array.from(dataStream).flatMap((token, startIndex) => token === T.SECTION_BREAK ? [{ startIndex, sectionId: `s-${startIndex}` }] : []),
                tables: [{ tableId: 'outer', startIndex: 0, endIndex: outer.length }, { tableId: 'inner', startIndex: outer.indexOf(inner), endIndex: outer.indexOf(inner) + inner.length }],
            },
            tableSource: { outer: table('outer', [200, 400]), inner: table('inner', texts.map(() => 150), nestedAutoFit) },
            documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, textStyle: { ff: 'Arial', fs: 11 }, pageSize: { width: 600, height: 800 }, marginLeft: 0, marginRight: 0 },
        });
        const skeleton = DocumentSkeleton.create(bed.viewModel, bed.sectionBreakConfig.localeService);
        try {
            skeleton.calculate();
            const outerTable = skeleton.getSkeletonData()!.pages[0].skeTables.get('outer')!;
            return outerTable.rows[0].cells[1].skeTables.get('inner')!;
        } finally {
            skeleton.dispose();
            bed.viewModel.dispose();
            bed.dataModel.dispose();
            canvas.mockRestore();
            Reflect.set(FontCache, '_context', previousContext);
            FontCache.invalidateMetrics(() => true);
            clearFontCreateConfigCache();
        }
    }

    function nestedColumnsOf(texts: string[]) {
        return nestedTableOf(texts).tableSource.tableColumns.map((column) => Math.round(column.size.width.v * 100) / 100);
    }

    it('marks a nested table that opens its cell, which Word draws past the cell\'s edges', () => {
        // The verifier's probes v7-annidate-larghe: Word clips a nested table wider than its cell only when a paragraph,
        // even an empty one, comes before it in the cell (18 cases out of 18); the cell's first block it draws whole.
        expect(nestedTableOf(['aa', long]).opensCell).toBe(true);
        expect(nestedTableOf(['aa', long], DataStreamTreeTokenType.PARAGRAPH).opensCell).toBeUndefined();
        expect(nestedTableOf(['aa', long], `prima${DataStreamTreeTokenType.PARAGRAPH}`).opensCell).toBeUndefined();
    });

    it('keeps the grid of an autofit table nested in a fixed table', () => {
        // Word 365 (probe v5-annidate of the fifth verification): two columns of 75 pt, with long text and with short
        expect(nestedColumnsOf(['aa', long])).toEqual([150, 150]);
        expect(nestedColumnsOf(['aa', 'bb'])).toEqual([150, 150]);
    });

    it('keeps the lengths of a table without a width of its own that fits', () => {
        const length = (value: number) => ({ type: 'dxa' as const, value });
        expect(columnsOf(['aa', 'bb'], { width: { type: 'auto', value: 0 }, columns: [length(100), length(150)] }, [300, 300]))
            .toEqual([100, 150]);
    });
});

describe('table auto fit', () => {
    it.each([
        { text: ['A', 'B'], grid: [96, 96], margins: [4, 8], expected: [11, 15] },
        { text: ['Reset padding'], grid: [107], margins: [9.6], expected: [64.6] },
        { text: ['Reset padding'], grid: [360], margins: [9.6], expected: [64.6] },
        { text: ['A', 'B'], grid: [24, 80], margins: [0, 0], expected: [24, 80], preferred: true },
        { text: ['Reset padding'], grid: [320], margins: [9.6], expected: [320], preferred: true },
        { text: ['Reset padding'], grid: [360], margins: [9.6], expected: [100.6], paragraphStyle: { indentStart: { v: 24 }, indentEnd: { v: 12 } } },
        { text: ['Reset padding'], grid: [360], margins: [9.6], expected: [80.6], paragraphStyle: { indentFirstLine: { v: 16 } } },
        { text: ['Reset padding'], grid: [360], margins: [9.6], expected: [1849 / 15], paragraphStyle: { indentStart: { v: 2, u: NumberUnitType.CHARACTER }, indentEnd: { v: 1, u: NumberUnitType.CHARACTER }, indentFirstLine: { v: 1, u: NumberUnitType.CHARACTER } } },
        { text: ['Reset padding'], grid: [360], margins: [9.6], expected: [80.6], paragraphStyle: { indentStart: { v: 0, u: NumberUnitType.CHARACTER, fallback: 16 } } },
        { text: ['Reset padding'], grid: [360], margins: [9.6], expected: [1409 / 15], paragraphStyle: { indentStart: { v: 2, u: NumberUnitType.CHARACTER }, hanging: { v: 1, u: NumberUnitType.CHARACTER } } },
    ])('fits automatic tables to content and paragraph indents while preserving public cell widths ($grid/$preferred/$paragraphStyle)', ({ text, grid, margins, expected, preferred, paragraphStyle }) => {
        const previousContext = Reflect.get(FontCache, '_context');
        Reflect.set(FontCache, '_context', null);
        FontCache.invalidateMetrics(() => true);
        const context = {
            font: '',
            measureText(this: CanvasRenderingContext2D, value: string) {
                const match = this.font.match(/([\d.]+)(pt|px)/);
                const points = match ? Number(match[1]) * (match[2] === 'px' ? 0.75 : 1) : 11;
                const ratio = points / 11;
                return { width: value.length * 4 * ratio, fontBoundingBoxAscent: 8 * ratio, fontBoundingBoxDescent: 2 * ratio, actualBoundingBoxAscent: 8 * ratio, actualBoundingBoxDescent: 2 * ratio };
            },
        } as unknown as CanvasRenderingContext2D;
        const canvas = vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(context);
        const T = DataStreamTreeTokenType;
        const row = `${T.TABLE_ROW_START}${text.map((value) => `${T.TABLE_CELL_START}${value}${T.PARAGRAPH}${T.SECTION_BREAK}${T.TABLE_CELL_END}`).join('')}${T.TABLE_ROW_END}`;
        const stream = `${T.TABLE_START}${row}${T.TABLE_END}`;
        const dataStream = `${stream}${T.PARAGRAPH}${T.SECTION_BREAK}`;
        const source: ITable = {
            tableId: 'content-fit',
            align: TableAlignmentType.START,
            indent: { v: 0 },
            textWrap: TableTextWrapType.NONE,
            position: { positionH: { relativeFrom: ObjectRelativeFromH.PAGE }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE } },
            dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
            size: { type: TableSizeType.UNSPECIFIED, width: { v: 0 } },
            cellMargin: { start: { v: 0 }, end: { v: 3 }, top: { v: 0 }, bottom: { v: 0 } },
            tableColumns: grid.map((width) => ({ size: { type: TableSizeType.SPECIFIED, width: { v: width } } })),
            tableRows: [{ trHeight: { hRule: TableRowHeightRule.AUTO, val: { v: 0 } }, tableCells: margins.map((start, column) => ({
                margin: { start: { v: start }, end: { v: 3 }, top: { v: 0 }, bottom: { v: 0 } },
                ...(preferred ? { size: { type: TableSizeType.SPECIFIED, width: { v: grid[column] } } } : {}),
            })) }],
        };
        const bed = createParagraphLayoutTestBed(stream, {
            body: {
                paragraphs: Array.from(dataStream).flatMap((token, startIndex) => token === T.PARAGRAPH ? [{ startIndex, paragraphId: `p-${startIndex}`, paragraphStyle: { lineSpacing: 16, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE, ...paragraphStyle } }] : []),
                sectionBreaks: Array.from(dataStream).flatMap((token, startIndex) => token === T.SECTION_BREAK ? [{ startIndex, sectionId: `s-${startIndex}` }] : []),
                tables: [{ tableId: source.tableId, startIndex: 0, endIndex: stream.length }],
            },
            tableSource: { [source.tableId]: source },
            documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, textStyle: { ff: 'Arial', fs: 11 } },
        });
        const before = JSON.stringify(bed.dataModel.getSnapshot());
        const skeleton = DocumentSkeleton.create(bed.viewModel, bed.sectionBreakConfig.localeService);
        try {
            skeleton.calculate();
            const actual = skeleton.getSkeletonData()!.pages[0].skeTables.get(source.tableId)!;
            expect(actual.tableSource.tableColumns.map((column) => column.size.width.v)).toEqual(expected);
            expect(actual.rows[0].cells.every((cell) => cell.sections[0].columns[0].lines.length === 1)).toBe(true);
            expect(JSON.stringify(bed.dataModel.getSnapshot())).toBe(before);
        } finally {
            skeleton.dispose();
            bed.viewModel.dispose();
            bed.dataModel.dispose();
            canvas.mockRestore();
            Reflect.set(FontCache, '_context', previousContext);
            FontCache.invalidateMetrics(() => true);
            clearFontCreateConfigCache();
        }
    });

    it.each([false, true].flatMap((incremental) => [false, true].flatMap((bordered) =>
        [0, 1, 6].map((spacing) => ({ incremental, bordered, spacing })))))('fits the preferred percentage column to unbreakable text and cell gaps ($incremental/$bordered/$spacing)', ({ incremental, bordered, spacing }) => {
        const previousContext = Reflect.get(FontCache, '_context');
        Reflect.set(FontCache, '_context', null);
        FontCache.invalidateMetrics(() => true);
        const context = {
            font: '',
            measureText(this: CanvasRenderingContext2D, text: string) {
                const match = this.font.match(/([\d.]+)(pt|px)/);
                const points = match ? Number(match[1]) * (match[2] === 'px' ? 0.75 : 1) : 11;
                const ratio = points / 11;
                return { width: text.length * 4 * ratio, fontBoundingBoxAscent: 8 * ratio, fontBoundingBoxDescent: 2 * ratio, actualBoundingBoxAscent: 8 * ratio, actualBoundingBoxDescent: 2 * ratio };
            },
        } as unknown as CanvasRenderingContext2D;
        const canvas = vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(context);
        const T = DataStreamTreeTokenType;
        const row = `${T.TABLE_ROW_START}${['[1] ', 'Reference study'].map((text) => `${T.TABLE_CELL_START}${text}${T.PARAGRAPH}${T.SECTION_BREAK}${T.TABLE_CELL_END}`).join('')}${T.TABLE_ROW_END}`;
        const stream = `${T.TABLE_START}${row}${T.TABLE_END}`;
        const dataStream = `${stream}${T.PARAGRAPH}${T.SECTION_BREAK}`;
        const border = { width: { v: bordered ? 2 : 0 }, color: { rgb: '#112233' } };
        const source: ITable = {
            tableId: 'fit',
            align: TableAlignmentType.START,
            indent: { v: 0 },
            textWrap: TableTextWrapType.NONE,
            position: { positionH: { relativeFrom: ObjectRelativeFromH.PAGE }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE } },
            dist: { distT: 0, distB: 0, distL: 0, distR: 0 },
            size: { type: TableSizeType.SPECIFIED, width: { v: 200 } },
            cellSpacing: { v: spacing },
            cellMargin: { start: { v: 1 }, end: { v: 1 }, top: { v: 1 }, bottom: { v: 1 } },
            tableColumns: [24, 176].map((width) => ({ size: { type: TableSizeType.SPECIFIED, width: { v: width } } })),
            tableRows: [{ trHeight: { hRule: TableRowHeightRule.AUTO, val: { v: 0 } }, tableCells: [0, 1].map((column) => ({
                borderTop: border,
                borderBottom: border,
                borderLeft: border,
                borderRight: border,
                tcWidthType: column === 0 ? 'pct' : 'auto',
                ...(column === 0 ? { size: { type: TableSizeType.SPECIFIED, width: { v: 50 / 15 } } } : {}),
            })) }],
        };
        const bed = createParagraphLayoutTestBed(stream, {
            body: {
                paragraphs: Array.from(dataStream).flatMap((token, startIndex) => token === T.PARAGRAPH ? [{ startIndex, paragraphId: `p-${startIndex}`, paragraphStyle: { lineSpacing: 16, spacingRule: SpacingRule.EXACT, snapToGrid: BooleanNumber.FALSE } }] : []),
                sectionBreaks: Array.from(dataStream).flatMap((token, startIndex) => token === T.SECTION_BREAK ? [{ startIndex, sectionId: `s-${startIndex}` }] : []),
                tables: [{ tableId: 'fit', startIndex: 0, endIndex: stream.length }],
            },
            tableSource: { fit: source },
            documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, textStyle: { ff: 'Arial', fs: 11 } },
        });
        const before = JSON.stringify(bed.dataModel.getSnapshot());
        try {
            const tableNode = bed.viewModel.getChildren()[0].children[0].children[0];
            const table = resolveTableLayoutColumns(source, 200, true, bed.dataModel.getSnapshot());
            const fitted = resolveTableAutoFitColumns(bed.ctx, bed.viewModel, tableNode, bed.sectionBreakConfig, table);
            const minimum = 14 + spacing * 3 + (bordered ? 4 : 0);
            expect(fitted.tableColumns[0].size.width.v).toBeCloseTo(minimum);
            expect(fitted.tableColumns[1].size.width.v).toBeCloseTo(200 - (bordered ? 2 : 0) - minimum);
            expect(JSON.stringify(bed.dataModel.getSnapshot())).toBe(before);
            for (const preference of [
                { tcWidthType: 'pct' as const, size: { type: TableSizeType.SPECIFIED, width: { v: 2500 / 15 } }, expected: bordered ? 99 : 100 },
                { tcWidthType: 'dxa' as const, size: { type: TableSizeType.SPECIFIED, width: { v: 80 } }, expected: 80 },
            ]) {
                const edited = { ...table, tableRows: table.tableRows.map((row) => ({ ...row, tableCells: row.tableCells.map((cell, column) => column === 0 ? { ...cell, tcWidthType: preference.tcWidthType, size: preference.size } : cell) })) };
                expect(resolveTableAutoFitColumns(bed.ctx, bed.viewModel, tableNode, bed.sectionBreakConfig, edited).tableColumns[0].size.width.v).toBeCloseTo(preference.expected);
            }
            const automatic = { ...table, tableRows: table.tableRows.map((row) => ({ ...row, tableCells: row.tableCells.map((cell) => ({ ...cell, tcWidthType: 'auto' as const })) })) };
            expect(resolveTableAutoFitColumns(bed.ctx, bed.viewModel, tableNode, bed.sectionBreakConfig, automatic).tableColumns[0].size.width.v).toBeCloseTo(minimum);
            const fixed = { ...table, layout: TableLayoutType.FIXED };
            expect(resolveTableAutoFitColumns(bed.ctx, bed.viewModel, tableNode, bed.sectionBreakConfig, fixed)).toBe(fixed);
            const skeleton = DocumentSkeleton.create(bed.viewModel, bed.sectionBreakConfig.localeService);
            try {
                if (incremental) {
                    const generation = skeleton.startIncrementalLayout();
                    let complete = false;
                    for (let count = 0; count < 1000 && !complete; count++) {
                        complete = skeleton.stepIncrementalLayout(generation, 1).complete;
                    }
                    expect(complete).toBe(true);
                } else {
                    skeleton.calculate();
                }
                const actual = skeleton.getSkeletonData()!.pages[0].skeTables.get('fit')!;
                expect(actual.tableSource.tableColumns[0].size.width.v).toBeCloseTo(minimum);
                expect(actual.width).toBeCloseTo(200);
                const lines = actual.rows[0].cells[0].sections[0].columns[0].lines;
                const numberedLines = lines.filter((line) => line.divides.some((divide) => divide.glyphGroup.some((glyph) => /\d/.test(glyph.raw))));
                expect(numberedLines).toHaveLength(1);
                expect(numberedLines[0].divides.flatMap((divide) => divide.glyphGroup.map((glyph) => glyph.raw)).join('').trim()).toBe('[1]');
                expect(JSON.stringify(bed.dataModel.getSnapshot())).toBe(before);
            } finally {
                skeleton.dispose();
            }
        } finally {
            bed.viewModel.dispose();
            bed.dataModel.dispose();
            canvas.mockRestore();
            Reflect.set(FontCache, '_context', previousContext);
            FontCache.invalidateMetrics(() => true);
            clearFontCreateConfigCache();
        }
    });
});
