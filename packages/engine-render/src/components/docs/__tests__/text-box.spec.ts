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

import type { IDocDrawingBase } from '@univerjs/core';
import { AlignTypeH, BooleanNumber, DocumentFlavor, ObjectRelativeFromH, ObjectRelativeFromV, PositionedObjectLayoutType, SpacingRule } from '@univerjs/core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { IDocumentSkeletonDrawing } from '../../../basics/i-document-skeleton-cached';
import type { UniverRenderingContext } from '../../../context';
import { setupRenderTestEnv } from '../../../__tests__/render-test-utils';
import { Canvas } from '../../../canvas';
import { Documents } from '../document';
import { createParagraphLayoutTestBed } from '../layout/block/paragraph/__tests__/create-paragraph-layout-test-bed';
import { DocumentSkeleton } from '../layout/doc-skeleton';
import { drawDocTextBoxShape, getDocTextBox, getDocTextBoxFlowStarts, getDocTextBoxFrame } from '../text-box';

function skeletonDrawing(origin: Partial<IDocDrawingBase>, aLeft = 10, aTop = 20): IDocumentSkeletonDrawing {
    return {
        drawingId: 'box',
        aLeft,
        aTop,
        width: 120,
        height: 40,
        angle: 0,
        initialState: true,
        drawingOrigin: origin as IDocDrawingBase,
        columnLeft: 0,
        isPageBreak: false,
        lineTop: 0,
        lineHeight: 0,
        blockAnchorTop: 0,
    };
}

const page = { pageWidth: 600, pageHeight: 800, marginLeft: 70, marginRight: 50 };

describe('text boxes', () => {
    it('reads the text box of a drawing, only when it has a content', () => {
        expect(getDocTextBox({ textBox: { body: { dataStream: 'A\r\n' } } } as unknown as IDocDrawingBase)?.body.dataStream).toBe('A\r\n');
        expect(getDocTextBox({ textBox: {} } as unknown as IDocDrawingBase)).toBeUndefined();
        expect(getDocTextBox({} as IDocDrawingBase)).toBeUndefined();
    });

    it('places the frame as the drawings are placed: from the page, its margins, or the segment', () => {
        const size = { width: 120, height: 40 };
        // offsets from the edges of the page, as the footer box of the Mosaico
        expect(getDocTextBoxFrame(skeletonDrawing({
            layoutType: PositionedObjectLayoutType.WRAP_NONE,
            docTransform: {
                size,
                positionH: { relativeFrom: ObjectRelativeFromH.PAGE, posOffset: 75 },
                positionV: { relativeFrom: ObjectRelativeFromV.PAGE, posOffset: 1045 / 1.5 },
                angle: 0,
            },
        } as Partial<IDocDrawingBase>), 700, 70, page)).toEqual({ left: 75, top: 1045 / 1.5, width: 120, height: 40 });
        // centred between the margins, below its paragraph
        expect(getDocTextBoxFrame(skeletonDrawing({
            layoutType: PositionedObjectLayoutType.WRAP_SQUARE,
            docTransform: {
                size,
                positionH: { relativeFrom: ObjectRelativeFromH.MARGIN, align: AlignTypeH.CENTER },
                positionV: { relativeFrom: ObjectRelativeFromV.PARAGRAPH, posOffset: 0 },
                angle: 0,
            },
        } as Partial<IDocDrawingBase>), 50, 70, page)).toEqual({ left: 70 + (480 - 120) / 2, top: 70, width: 120, height: 40 });
        // in line: where the line put it in the segment
        expect(getDocTextBoxFrame(skeletonDrawing({
            layoutType: PositionedObjectLayoutType.INLINE,
            docTransform: { size, positionH: { relativeFrom: ObjectRelativeFromH.PAGE, posOffset: 300 }, positionV: { relativeFrom: ObjectRelativeFromV.PAGE, posOffset: 300 }, angle: 0 },
        } as Partial<IDocDrawingBase>), 50, 70, page)).toEqual({ left: 80, top: 70, width: 120, height: 40 });
    });

    it('draws the shape with its line centred on the frame, the corners of a rounded rectangle as arcs', () => {
        const ctx = {
            save: vi.fn(),
            restore: vi.fn(),
            beginPath: vi.fn(),
            closePath: vi.fn(),
            rect: vi.fn(),
            moveTo: vi.fn(),
            lineTo: vi.fn(),
            bezierCurveTo: vi.fn(),
            fill: vi.fn(),
            stroke: vi.fn(),
            setLineDash: vi.fn(),
        } as unknown as UniverRenderingContext & Record<string, ReturnType<typeof vi.fn>>;
        const frame = { left: 10, top: 20, width: 300, height: 60 };
        drawDocTextBoxShape(ctx, { geometry: 'rect', line: { color: '#196B24', width: 2.5, dash: 'dash' } }, frame);
        expect(ctx.rect).toHaveBeenCalledWith(10, 20, 300, 60);
        expect(ctx.stroke).toHaveBeenCalledTimes(1);
        expect(ctx.fill).not.toHaveBeenCalled();
        expect(ctx.lineWidth).toBe(2.5);
        expect(ctx.setLineDash).toHaveBeenCalledWith([10, 7.5]);

        drawDocTextBoxShape(ctx, { geometry: 'flowChartAlternateProcess', fill: { color: '#FFFFFF' }, line: { color: '#196B24', width: 2.5 } }, frame);
        expect(ctx.bezierCurveTo).toHaveBeenCalledTimes(4);
        // the corner radius of the alternate process: a sixth of the shortest side
        expect(ctx.moveTo).toHaveBeenCalledWith(10 + 10, 20);
        expect(ctx.fill).toHaveBeenCalledTimes(1);

        (ctx.stroke as unknown as ReturnType<typeof vi.fn>).mockClear();
        drawDocTextBoxShape(ctx, { geometry: 'rect' }, frame);
        expect(ctx.stroke).not.toHaveBeenCalled();
    });
});

describe('linked text boxes', () => {
    it('starts each box of a chain with the first line that does not fit in the boxes before it', () => {
        const lines: Array<[number, number]> = [[0, 10], [10, 20], [20, 30], [30, 40]];
        expect(getDocTextBoxFlowStarts(lines, [25, 30])).toEqual([0, 20]);
        expect(getDocTextBoxFlowStarts(lines, [25, 15, 100])).toEqual([0, 20, 30]);
        // a box the text does not reach shows nothing
        expect(getDocTextBoxFlowStarts(lines, [100, 50])).toEqual([0, Number.POSITIVE_INFINITY]);
    });
});

describe('text boxes on the page', () => {
    let env: { restore: () => void };
    let canvas: Canvas;

    beforeEach(() => {
        env = setupRenderTestEnv();
        canvas = new Canvas({ width: 800, height: 900, pixelRatio: 1 });
    });

    afterEach(() => {
        canvas.dispose();
        env.restore();
        vi.restoreAllMocks();
    });

    /** Draws a page with a text box anchored at the given place: the glyphs where they land, the rectangles. */
    function drawBox(verticalAlign: 'top' | 'middle' | 'bottom') {
        const box = {
            drawingId: 'box',
            unitId: 'test-doc',
            subUnitId: 'test-doc',
            drawingType: 0,
            imageSourceType: 'BASE64',
            source: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciLz4=',
            transform: { width: 200, height: 80, left: 0, top: 0, angle: 0 },
            docTransform: {
                size: { width: 200, height: 80 },
                positionH: { relativeFrom: ObjectRelativeFromH.PAGE, posOffset: 100 },
                positionV: { relativeFrom: ObjectRelativeFromV.PAGE, posOffset: 300 },
                angle: 0,
            },
            layoutType: PositionedObjectLayoutType.WRAP_NONE,
            behindDoc: 0,
            distT: 0,
            distB: 0,
            distL: 0,
            distR: 0,
            textBox: {
                body: {
                    dataStream: 'Casella\r\n',
                    textRuns: [{ st: 0, ed: 9, ts: { fs: 12 } }],
                    paragraphs: [{ startIndex: 7, paragraphId: 'box-paragraph' }],
                    sectionBreaks: [{ startIndex: 8, sectionId: 'box-section' }],
                },
                insets: { left: 10, top: 6, right: 10, bottom: 6 },
                verticalAlign,
                shape: { geometry: 'rect', line: { color: '#196B24', width: 2 } },
            },
        };
        const bed = createParagraphLayoutTestBed('A\bB', {
            documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, pageSize: { width: 600, height: 800 } },
            body: {
                customBlocks: [{ startIndex: 1, blockId: 'box' }],
                textRuns: [{ st: 0, ed: 5, ts: { fs: 12 } }],
                paragraphs: [{ startIndex: 3, paragraphId: 'anchor' }],
                sectionBreaks: [{ startIndex: 4, sectionId: 'section' }],
            },
            drawings: { box },
            drawingsOrder: ['box'],
        });
        const skeleton = DocumentSkeleton.create(bed.viewModel, bed.ctx.docsConfig.localeService);
        const documents = bed.injector.createInstance(Documents, `text-box-${verticalAlign}`, skeleton, { pageMarginLeft: 0, pageMarginTop: 0 });
        try {
            skeleton.calculate();
            const ctx = canvas.getContext();
            // where each glyph lands on the page: its origin through the transform of the context
            const glyphs: Array<{ text: string; x: number; y: number }> = [];
            const fillText = ctx.fillText.bind(ctx);
            const text = vi.spyOn(ctx, 'fillText').mockImplementation((value: string, x: number, y: number) => {
                const m = ctx.getTransform();
                glyphs.push({ text: String(value), x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f });
                fillText(value, x, y);
            });
            const rect = vi.spyOn(ctx, 'rect');
            documents.draw(ctx);
            const rects = rect.mock.calls.map((call) => call.slice(0, 4));
            text.mockRestore();
            rect.mockRestore();
            return { glyphs, rects };
        } finally {
            documents.dispose();
            skeleton.dispose();
            bed.viewModel.dispose();
            bed.dataModel.dispose();
        }
    }

    it('lays the content out in the frame of the drawing, within its insets and at its vertical anchor', () => {
        const top = drawBox('top');
        expect(top.glyphs.map((glyph) => glyph.text).join('')).toContain('Casella');
        // the shape where Word puts the frame, the text cut at the bottom of the room inside the insets
        expect(top.rects).toContainEqual([100, 300, 200, 80]);
        expect(top.rects).toContainEqual([100, 306, 200, 68]);
        const at = (drawn: typeof top) => drawn.glyphs.find((glyph) => glyph.text.startsWith('C'))!;
        expect(at(top).x).toBeGreaterThanOrEqual(110 - 0.01);
        expect(at(top).y).toBeGreaterThanOrEqual(306);
        // middle and bottom move the content by a half and by all of what the room leaves under it
        const middle = at(drawBox('middle')).y - at(top).y;
        const bottom = at(drawBox('bottom')).y - at(top).y;
        expect(middle).toBeGreaterThan(0);
        expect(bottom).toBeCloseTo(2 * middle, 5);
        expect(at(top).y + bottom).toBeLessThanOrEqual(374);
    });

    it('flows the text of the first box of a chain into the next box, as Word', () => {
        // C1 of the third independent verification: the second box of a chain has no content of its own, and the
        // text that does not fit in the first was lost, cut by its frame.
        const frameOf = (left: number, top: number, height: number) => ({
            unitId: 'test-doc',
            subUnitId: 'test-doc',
            drawingType: 0,
            imageSourceType: 'BASE64',
            source: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciLz4=',
            transform: { width: 200, height, left: 0, top: 0, angle: 0 },
            docTransform: {
                size: { width: 200, height },
                positionH: { relativeFrom: ObjectRelativeFromH.PAGE, posOffset: left },
                positionV: { relativeFrom: ObjectRelativeFromV.PAGE, posOffset: top },
                angle: 0,
            },
            layoutType: PositionedObjectLayoutType.WRAP_NONE,
            behindDoc: 0,
            distT: 0,
            distB: 0,
            distL: 0,
            distR: 0,
        });
        const insets = { left: 10, top: 6, right: 10, bottom: 6 };
        const text = 'Uno\rDue\rTre\rQuattro\r\n';
        // lines 15 px high, whatever the font metrics of the test canvas
        const paragraphs = [3, 7, 11, 19].map((startIndex, index) => ({
            startIndex,
            paragraphId: `p${index}`,
            paragraphStyle: { spacingRule: SpacingRule.EXACT, lineSpacing: 15, snapToGrid: BooleanNumber.FALSE },
        }));
        const flow = { head: 'first', rooms: [40, 100], width: 200 };
        const first = {
            ...frameOf(100, 100, 52),
            drawingId: 'first',
            textBox: {
                body: { dataStream: text, textRuns: [{ st: 0, ed: 21, ts: { fs: 12 } }], paragraphs, sectionBreaks: [{ startIndex: 20, sectionId: 's' }] },
                insets,
                flow: { ...flow, order: 0 },
            },
        };
        const second = {
            ...frameOf(350, 300, 112),
            drawingId: 'second',
            textBox: {
                body: { dataStream: '\r\n', paragraphs: [{ startIndex: 0, paragraphId: 'e' }], sectionBreaks: [{ startIndex: 1, sectionId: 'e' }] },
                insets,
                flow: { ...flow, order: 1 },
            },
        };
        const bed = createParagraphLayoutTestBed('A\b\bB', {
            documentStyle: { documentFlavor: DocumentFlavor.TRADITIONAL, pageSize: { width: 600, height: 800 } },
            body: {
                customBlocks: [{ startIndex: 1, blockId: 'first' }, { startIndex: 2, blockId: 'second' }],
                textRuns: [{ st: 0, ed: 6, ts: { fs: 12 } }],
                paragraphs: [{ startIndex: 4, paragraphId: 'anchor' }],
                sectionBreaks: [{ startIndex: 5, sectionId: 'section' }],
            },
            drawings: { first, second },
            drawingsOrder: ['first', 'second'],
        });
        const skeleton = DocumentSkeleton.create(bed.viewModel, bed.ctx.docsConfig.localeService);
        const documents = bed.injector.createInstance(Documents, 'text-box-chain', skeleton, { pageMarginLeft: 0, pageMarginTop: 0 });
        try {
            skeleton.calculate();
            const ctx = canvas.getContext();
            const glyphs: Array<{ text: string; x: number; y: number }> = [];
            const fillText = ctx.fillText.bind(ctx);
            vi.spyOn(ctx, 'fillText').mockImplementation((value: string, x: number, y: number) => {
                const m = ctx.getTransform();
                glyphs.push({ text: String(value), x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f });
                fillText(value, x, y);
            });
            const rect = vi.spyOn(ctx, 'rect');
            documents.draw(ctx);
            const rects = rect.mock.calls.map((call) => call.slice(0, 4) as number[]);
            const inSecond = glyphs.filter((glyph) => glyph.x >= 350 && glyph.y >= 300);
            const inFirst = glyphs.filter((glyph) => glyph.x < 350 && glyph.y < 300);
            // the first box shows the lines that fit in its room, cut where the next line starts
            const firstClip = rects.find((r) => r[0] === 100 && r[1] === 106)!;
            expect(firstClip[3]).toBeLessThanOrEqual(40);
            expect(inFirst.filter((glyph) => glyph.text === 'U').length).toBe(1);
            // the second box goes on with the text, from the top of its room
            const t = inSecond.find((glyph) => glyph.text === 'T')!;
            const q = inSecond.find((glyph) => glyph.text === 'Q')!;
            expect(t.x).toBeGreaterThanOrEqual(360 - 0.01);
            expect(t.y).toBeGreaterThan(306);
            expect(t.y).toBeLessThan(306 + 20);
            expect(q.y).toBeGreaterThan(t.y);
            // the lines of the first box are drawn above the room of the second, where its clip leaves them out
            const secondClip = rects.find((r) => r[0] === 350 && r[1] === 306)!;
            expect(secondClip[3]).toBe(100);
            expect(inSecond.some((glyph) => glyph.text === 'U' && glyph.y >= 306)).toBe(false);
        } finally {
            documents.dispose();
            skeleton.dispose();
            bed.viewModel.dispose();
            bed.dataModel.dispose();
        }
    });
});
