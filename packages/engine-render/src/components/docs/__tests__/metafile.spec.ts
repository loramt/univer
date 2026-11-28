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

import { describe, expect, it, vi } from 'vitest';
import type { UniverRenderingContext } from '../../../context';
import { drawEmf, isEmf, parseEmf } from '../metafile';

/** An EMF as Excel writes a table pasted as a picture: GDI records in device units of a 96 dpi screen. */
function excelLikeEmf(): Uint8Array {
    const records: number[][] = [];
    const u32 = (value: number) => [value & 0xFF, (value >>> 8) & 0xFF, (value >>> 16) & 0xFF, (value >>> 24) & 0xFF];
    const i32 = (value: number) => u32(value >>> 0);
    const record = (type: number, body: number[]) => records.push([...u32(type), ...u32(8 + body.length), ...body]);
    // header: bounds, frame of 100 x 50 px at 96 dpi (2645.83 x 1322.92 in 0.01 mm), device 1000 x 500 px, 264.58 x 132.29 mm
    const header = [
        ...i32(0),
        ...i32(0),
        ...i32(99),
        ...i32(49),
        ...i32(0),
        ...i32(0),
        ...i32(2646),
        ...i32(1323),
        ...u32(0x464D4520),
        ...u32(0x10000),
        ...u32(0),
        ...u32(0),
        ...u32(0),
        ...u32(0),
        ...u32(0),
        ...u32(0),
        ...i32(1000),
        ...i32(500),
        ...i32(265),
        ...i32(132),
        ...u32(0),
        ...u32(0),
        ...u32(0),
        ...u32(264583),
        ...u32(132292),
    ];
    record(1, header);
    // a green fill of the first row: a pattern copy with the brush
    record(39, [...u32(1), ...u32(0), ...u32(0x0050B000), ...u32(0)]);
    record(37, u32(1));
    record(76, [...i32(0), ...i32(0), ...i32(99), ...i32(19), ...i32(0), ...i32(0), ...i32(100), ...i32(20), ...u32(0x00F00021), ...i32(0), ...i32(0), ...new Array(24).fill(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0)]);
    // a black line under it
    record(38, [...u32(2), ...u32(0), ...i32(1), ...i32(0), ...u32(0)]);
    record(37, u32(2));
    record(27, [...i32(0), ...i32(20)]);
    record(54, [...i32(100), ...i32(20)]);
    // "Ab" in bold Calibri of 15 px (11 pt), from its top, with the advances of Excel
    const face = [...'Calibri'].flatMap((char) => [char.charCodeAt(0), 0]);
    record(82, [...u32(3), ...i32(-15), ...i32(0), ...i32(0), ...i32(0), ...i32(700), 0, 0, 0, 0, 0, 0, 0, 0, ...face, ...new Array(64 - face.length).fill(0), ...new Array(256).fill(0)]);
    record(37, u32(3));
    record(24, u32(0x000000FF));
    const text = [...'Ab'].flatMap((char) => [char.charCodeAt(0), 0]);
    // EMR_EXTTEXTOUTW: bounds, graphics mode, scales, then EMRTEXT with the string at 76 and the advances at 80
    record(84, [...i32(0), ...i32(0), ...i32(0), ...i32(0), ...u32(1), ...u32(0), ...u32(0), ...i32(4), ...i32(25), ...u32(2), ...u32(76), ...u32(0), ...i32(0), ...i32(0), ...i32(-1), ...i32(-1), ...u32(80), ...text, ...i32(9), ...i32(8)]);
    record(14, [...u32(0), ...u32(16), ...u32(20)]);
    const bytes = new Uint8Array(records.flat());
    // the size of the header record covers the extension up to the micrometres
    return bytes;
}

/** An EMF as PowerPoint writes a rounded rectangle: a path of moves, Bézier curves and lines, filled with a brush. */
function pathEmf(): Uint8Array {
    const records: number[][] = [];
    const u32 = (value: number) => [value & 0xFF, (value >>> 8) & 0xFF, (value >>> 16) & 0xFF, (value >>> 24) & 0xFF];
    const i32 = (value: number) => u32(value >>> 0);
    const i16 = (value: number) => [value & 0xFF, (value >> 8) & 0xFF];
    const record = (type: number, body: number[]) => records.push([...u32(type), ...u32(8 + body.length), ...body]);
    record(1, [...i32(0), ...i32(0), ...i32(99), ...i32(49), ...i32(0), ...i32(0), ...i32(2646), ...i32(1323), ...u32(0x464D4520), ...u32(0x10000), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...i32(1000), ...i32(500), ...i32(265), ...i32(132), ...u32(0), ...u32(0), ...u32(0), ...u32(264583), ...u32(132292)]);
    record(39, [...u32(1), ...u32(0), ...u32(0x00684F1F), ...u32(0)]);
    record(37, u32(1));
    const bezierTo = (points: number[]) => record(88, [...i32(0), ...i32(0), ...i32(0), ...i32(0), ...u32(points.length / 2), ...points.flatMap(i16)]);
    record(59, []);
    record(27, [...i32(10), ...i32(20)]);
    bezierTo([10, 15, 15, 10, 20, 10]);
    record(54, [...i32(80), ...i32(10)]);
    bezierTo([85, 10, 90, 15, 90, 20]);
    record(54, [...i32(90), ...i32(40)]);
    record(54, [...i32(10), ...i32(40)]);
    record(61, []);
    record(60, []);
    record(62, [...i32(0), ...i32(0), ...i32(99), ...i32(49)]);
    record(14, [...u32(0), ...u32(16), ...u32(20)]);
    return new Uint8Array(records.flat());
}

/** The table of excelLikeEmf with its text in a font turned by 90 degrees, as Excel writes a turned cell. */
function turnedTextEmf(): Uint8Array {
    const bytes = excelLikeEmf();
    const view = new DataView(bytes.buffer);
    // the escapement of the font record (EMR_EXTCREATEFONTINDIRECTW, type 82): 900 tenths of a degree
    for (let offset = 0; offset + 8 <= bytes.length;) {
        const type = view.getUint32(offset, true);
        const size = view.getUint32(offset + 4, true);
        if (type === 82) {
            view.setInt32(offset + 20, 900, true);
        }
        offset += size;
    }
    return bytes;
}

/** A dotted line as Excel draws its dotted borders: a pattern copy, one pixel high, with a pattern of dots. */
function dottedEmf(): Uint8Array {
    const records: number[][] = [];
    const u32 = (value: number) => [value & 0xFF, (value >>> 8) & 0xFF, (value >>> 16) & 0xFF, (value >>> 24) & 0xFF];
    const i32 = (value: number) => u32(value >>> 0);
    const record = (type: number, body: number[]) => records.push([...u32(type), ...u32(8 + body.length), ...body]);
    record(1, [...i32(0), ...i32(0), ...i32(99), ...i32(49), ...i32(0), ...i32(0), ...i32(2646), ...i32(1323), ...u32(0x464D4520), ...u32(0x10000), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...i32(1000), ...i32(500), ...i32(265), ...i32(132), ...u32(0), ...u32(0), ...u32(0), ...u32(264583), ...u32(132292)]);
    // EMR_CREATEDIBPATTERNBRUSHPT: an 8 x 8 bitmap of one bit, black and white, every other pixel set
    const info = [...u32(40), ...i32(8), ...i32(8), 1, 0, 1, 0, ...u32(0), ...u32(32), ...i32(0), ...i32(0), ...u32(2), ...u32(0)];
    const palette = [0, 0, 0, 0, 255, 255, 255, 0];
    const bits = new Array(8).fill([0x55, 0, 0, 0]).flat();
    record(94, [...u32(1), ...u32(0), ...u32(32), ...u32(48), ...u32(80), ...u32(32), ...info, ...palette, ...bits]);
    record(37, u32(1));
    // EMR_BITBLT PATCOPY of 8 x 1 pixels, without a source
    record(76, [...i32(0), ...i32(0), ...i32(7), ...i32(0), ...i32(0), ...i32(10), ...i32(8), ...i32(1), ...u32(0x00F00021), ...i32(0), ...i32(0), ...new Array(24).fill(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0), ...u32(0)]);
    record(14, [...u32(0), ...u32(16), ...u32(20)]);
    return new Uint8Array(records.flat());
}

function recordingContext() {
    const calls: Array<{ op: string; args: unknown[]; fillStyle?: string; strokeStyle?: string; font?: string; lineWidth?: number }> = [];
    const ctx: Record<string, unknown> = { font: '', fillStyle: '', strokeStyle: '', lineWidth: 1 };
    for (const op of ['save', 'restore', 'beginPath', 'rect', 'clip', 'moveTo', 'lineTo', 'stroke', 'fill', 'fillRect', 'fillText', 'closePath', 'setLineDash', 'bezierCurveTo', 'ellipse', 'drawImage', 'translate', 'rotate']) {
        ctx[op] = vi.fn((...args: unknown[]) => calls.push({ op, args, fillStyle: ctx.fillStyle as string, strokeStyle: ctx.strokeStyle as string, font: ctx.font as string, lineWidth: ctx.lineWidth as number }));
    }
    ctx.measureText = vi.fn(() => ({ width: 10, fontBoundingBoxAscent: 12, fontBoundingBoxDescent: 3 }));
    return { ctx: ctx as unknown as UniverRenderingContext, calls };
}

describe('Windows metafiles', () => {
    it('reads the frame of the picture in device units', () => {
        const bytes = excelLikeEmf();
        expect(isEmf(bytes)).toBe(true);
        expect(isEmf(new Uint8Array(100))).toBe(false);
        const emf = parseEmf(bytes)!;
        expect(emf.frame.left).toBeCloseTo(0);
        expect(emf.frame.right).toBeCloseTo(100, 0);
        expect(emf.frame.bottom).toBeCloseTo(50, 0);
        expect(emf.records.at(-1)?.type).toBe(14);
    });

    it('plays fills, lines and text at the positions of the file, in the frame of the drawing', () => {
        const emf = parseEmf(excelLikeEmf())!;
        const { ctx, calls } = recordingContext();
        // the picture at twice its size, from (10, 20) of the page
        drawEmf(ctx, emf, { left: 10, top: 20, width: 200, height: 100 });
        const scale = 200 / (emf.frame.right - emf.frame.left);
        // the fill of the first row, with the colour of the brush (COLORREF is red, green, blue from the low byte)
        const fill = calls.find((call) => call.op === 'fillRect' && call.fillStyle === 'rgb(0,176,80)');
        expect(fill?.args[0]).toBeCloseTo(10);
        expect(fill?.args[2]).toBeCloseTo(100 * scale);
        // the line, a pixel of the reference device wide
        const stroke = calls.find((call) => call.op === 'stroke');
        expect(stroke?.lineWidth).toBeCloseTo(scale);
        // the text, glyph by glyph at the advances of the file, its top at the reference point
        const glyphs = calls.filter((call) => call.op === 'fillText');
        expect(glyphs.map((call) => call.args[0])).toEqual(['A', 'b']);
        expect(glyphs[0].args[1]).toBeCloseTo(10 + 4 * scale);
        expect(glyphs[1].args[1]).toBeCloseTo(10 + (4 + 9) * scale);
        expect(glyphs[0].args[2]).toBeCloseTo(20 + 25 * scale + 12);
        expect(glyphs[0].fillStyle).toBe('rgb(255,0,0)');
        const scaleY = 100 / (emf.frame.bottom - emf.frame.top);
        expect(glyphs[0].font).toMatch(/^bold [\d.]+px "Calibri"$/);
        expect(Number.parseFloat(glyphs[0].font!.slice(5))).toBeCloseTo(15 * scaleY, 6);
    });

    it('copies a pattern brush pixel by pixel, as the dotted borders of Excel', () => {
        // F3 of the third independent verification: the dotted borders came out as full grey lines
        const emf = parseEmf(dottedEmf())!;
        const { ctx, calls } = recordingContext();
        drawEmf(ctx, emf, { left: 0, top: 0, width: 100, height: 50 });
        const dots = calls.filter((call) => call.op === 'fillRect');
        expect(dots.map((call) => call.fillStyle)).toEqual(['rgb(0,0,0)', 'rgb(255,255,255)', 'rgb(0,0,0)', 'rgb(255,255,255)', 'rgb(0,0,0)', 'rgb(255,255,255)', 'rgb(0,0,0)', 'rgb(255,255,255)']);
        const scale = 100 / (emf.frame.right - emf.frame.left);
        expect(dots[1].args[0]).toBeCloseTo(1 * scale);
        expect(dots[1].args[2]).toBeCloseTo(scale);
    });

    it('writes the text of a turned font along its turned baseline', () => {
        // F2 of the third independent verification: the text of Excel turned by 90 degrees in a cell came out flat
        const emf = parseEmf(turnedTextEmf())!;
        const { ctx, calls } = recordingContext();
        drawEmf(ctx, emf, { left: 10, top: 20, width: 200, height: 100 });
        const turn = calls.find((call) => call.op === 'rotate');
        expect(turn?.args[0]).toBeCloseTo(-Math.PI / 2);
        const scale = 200 / (emf.frame.right - emf.frame.left);
        const at = calls.findIndex((call) => call.op === 'rotate');
        const moved = calls.slice(0, at).filter((call) => call.op === 'translate').at(-1);
        expect(moved?.args[0]).toBeCloseTo(10 + 4 * scale);
        expect(moved?.args[1]).toBeCloseTo(20 + 25 * scale);
        const glyphs = calls.filter((call) => call.op === 'fillText');
        expect(glyphs.map((call) => call.args[1])).toEqual([0, 9 * scale]);
    });

    it('fills a path of moves, Bézier curves and lines as one figure, as GDI traces it', () => {
        // F1 of the third independent verification: each POLYBEZIERTO of the path of a rounded rectangle of PowerPoint
        // started a figure of its own, and the fill left a skewed white shape inside the blue one
        const emf = parseEmf(pathEmf())!;
        const { ctx, calls } = recordingContext();
        drawEmf(ctx, emf, { left: 0, top: 0, width: 100, height: 50 });
        const fill = calls.findIndex((call) => call.op === 'fill');
        const begin = calls.slice(0, fill).map((call) => call.op).lastIndexOf('beginPath');
        const traced = calls.slice(begin, fill).map((call) => call.op);
        expect(traced.filter((op) => op === 'moveTo')).toHaveLength(1);
        expect(traced.filter((op) => op === 'bezierCurveTo')).toHaveLength(2);
        expect(traced.filter((op) => op === 'lineTo')).toHaveLength(3);
        expect(calls[fill].fillStyle).toBe('rgb(31,79,104)');
    });
});
