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

import type { UniverRenderingContext } from '../../context';

/**
 * Windows enhanced metafiles (EMF), as Word draws them: the records of the GDI drawing played on the canvas, with
 * the text as text at the positions of its characters (EMR_EXTTEXTOUTW and its advances), lines, rectangles,
 * polygons, paths, clipping and bitmaps. The tables of Excel pasted as pictures use a small part of GDI; what is
 * not played (EMF+ comments, palettes, raster operations other than copies) is left out, the GDI records of dual
 * files carry the whole picture.
 */

/** Where the picture goes on the page, in the coordinates of the context. */
export interface IMetafileFrame {
    left: number;
    top: number;
    width: number;
    height: number;
}

/** A parsed EMF: its header and the view of its records. */
export interface IMetafile {
    /** The picture frame, in device units of the reference device (the frame of the header, 0.01 mm, converted). */
    frame: { left: number; top: number; right: number; bottom: number };
    /** Device units of the reference device per millimetre, for the metric mapping modes. */
    perMillimeter: { x: number; y: number };
    view: DataView;
    records: Array<{ type: number; offset: number; size: number }>;
}

const EMR_HEADER = 1;
const EMR_EOF = 14;

/** Whether the bytes are an enhanced metafile (EMR_HEADER with the signature " EMF"). */
export function isEmf(bytes: Uint8Array): boolean {
    if (bytes.length < 88) {
        return false;
    }
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return view.getUint32(0, true) === EMR_HEADER && view.getUint32(40, true) === 0x464D4520;
}

/** The records and the frame of an EMF, or undefined when the bytes are not one. */
export function parseEmf(bytes: Uint8Array): IMetafile | undefined {
    if (!isEmf(bytes)) {
        return undefined;
    }
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const headerSize = view.getUint32(4, true);
    const frameLeft = view.getInt32(24, true);
    const frameTop = view.getInt32(28, true);
    const frameRight = view.getInt32(32, true);
    const frameBottom = view.getInt32(36, true);
    const deviceX = view.getInt32(72, true);
    const deviceY = view.getInt32(76, true);
    let millimetersX = view.getInt32(80, true);
    let millimetersY = view.getInt32(84, true);
    // the size of the reference device in micrometres, when the header has it, is more precise than in millimetres
    if (headerSize >= 108 && view.byteLength >= 108) {
        const micrometersX = view.getUint32(100, true);
        const micrometersY = view.getUint32(104, true);
        if (micrometersX > 0 && micrometersY > 0) {
            millimetersX = micrometersX / 1000;
            millimetersY = micrometersY / 1000;
        }
    }
    if (deviceX <= 0 || deviceY <= 0 || millimetersX <= 0 || millimetersY <= 0) {
        return undefined;
    }
    // device pixels per 0.01 mm
    const perX = deviceX / (millimetersX * 100);
    const perY = deviceY / (millimetersY * 100);
    const records: IMetafile['records'] = [];
    for (let offset = 0; offset + 8 <= view.byteLength;) {
        const type = view.getUint32(offset, true);
        const size = view.getUint32(offset + 4, true);
        if (size < 8 || offset + size > view.byteLength) {
            break;
        }
        records.push({ type, offset, size });
        offset += size;
        if (type === EMR_EOF) {
            break;
        }
    }
    return {
        frame: {
            left: frameLeft * perX,
            top: frameTop * perY,
            right: frameRight * perX,
            bottom: frameBottom * perY,
        },
        perMillimeter: { x: perX * 100, y: perY * 100 },
        view,
        records,
    };
}

interface IPen { color: string; width: number; style: number; cosmetic: boolean }
/**
 * The pattern of a brush (EMR_CREATEDIBPATTERNBRUSHPT, EMR_CREATEMONOBRUSH), from the top row: the bits (0 or 1) of a
 * monochrome one with its two colours, or the colour of each pixel.
 */
interface IBrushPattern { width: number; height: number; bits: Uint8Array; colors: [string, string] | null; pixels?: string[] }
interface IBrush { color: string | null; pattern?: IBrushPattern }
interface IFont { height: number; weight: number; italic: boolean; underline: boolean; strikeOut: boolean; escapement: number; face: string }
type Rect = [number, number, number, number];
type Matrix = [number, number, number, number, number, number];

interface IState {
    mapMode: number;
    windowOrg: [number, number];
    windowExt: [number, number];
    viewportOrg: [number, number];
    viewportExt: [number, number];
    world: Matrix;
    pen: IPen;
    brush: IBrush;
    font: IFont;
    textColor: string;
    bkColor: string;
    bkMode: number;
    textAlign: number;
    polyFillMode: number;
    position: [number, number];
    /** The clip region in device units, as rectangles, or null for none. */
    clip: Rect[] | null;
}

const BLACK = 'rgb(0,0,0)';
const STOCK: Record<number, IPen | IBrush | IFont> = {};

function colorRef(value: number): string {
    return `rgb(${value & 0xFF},${(value >>> 8) & 0xFF},${(value >>> 16) & 0xFF})`;
}

function stockObject(index: number): IPen | IBrush | IFont | undefined {
    if (STOCK[index]) {
        return STOCK[index];
    }
    const brush = (color: string | null): IBrush => ({ color });
    const pen = (color: string | null): IPen => ({ color: color ?? BLACK, width: 0, style: color === null ? 5 : 0, cosmetic: true });
    const objects: Record<number, IPen | IBrush | IFont> = {
        0: brush('rgb(255,255,255)'),
        1: brush('rgb(192,192,192)'),
        2: brush('rgb(128,128,128)'),
        3: brush('rgb(64,64,64)'),
        4: brush(BLACK),
        5: brush(null),
        6: pen('rgb(255,255,255)'),
        7: pen(BLACK),
        8: pen(null),
        10: { height: -16, weight: 400, italic: false, underline: false, strikeOut: false, escapement: 0, face: 'Courier New' },
        11: { height: -16, weight: 400, italic: false, underline: false, strikeOut: false, escapement: 0, face: 'Arial' },
        12: { height: -16, weight: 400, italic: false, underline: false, strikeOut: false, escapement: 0, face: 'Arial' },
        13: { height: -16, weight: 700, italic: false, underline: false, strikeOut: false, escapement: 0, face: 'Arial' },
        16: { height: -16, weight: 400, italic: false, underline: false, strikeOut: false, escapement: 0, face: 'Arial' },
        17: { height: -16, weight: 400, italic: false, underline: false, strikeOut: false, escapement: 0, face: 'Arial' },
    };
    if (objects[index]) {
        STOCK[index] = objects[index];
    }
    return objects[index];
}

function multiply(a: Matrix, b: Matrix): Matrix {
    // a then b: x' = b(a(x))
    return [
        a[0] * b[0] + a[1] * b[2],
        a[0] * b[1] + a[1] * b[3],
        a[2] * b[0] + a[3] * b[2],
        a[2] * b[1] + a[3] * b[3],
        a[4] * b[0] + a[5] * b[2] + b[4],
        a[4] * b[1] + a[5] * b[3] + b[5],
    ];
}

function intersect(region: Rect[] | null, rect: Rect): Rect[] {
    const base = region ?? [rect];
    const out: Rect[] = [];
    for (const [l, t, r, b] of base) {
        const left = Math.max(l, rect[0]);
        const top = Math.max(t, rect[1]);
        const right = Math.min(r, rect[2]);
        const bottom = Math.min(b, rect[3]);
        if (right > left && bottom > top) {
            out.push([left, top, right, bottom]);
        }
    }
    return out;
}

function subtract(region: Rect[] | null, rect: Rect, whole: Rect): Rect[] {
    const out: Rect[] = [];
    for (const [l, t, r, b] of region ?? [whole]) {
        const left = Math.max(l, rect[0]);
        const top = Math.max(t, rect[1]);
        const right = Math.min(r, rect[2]);
        const bottom = Math.min(b, rect[3]);
        if (right <= left || bottom <= top) {
            out.push([l, t, r, b]);
            continue;
        }
        if (top > t) out.push([l, t, r, top]);
        if (bottom < b) out.push([l, bottom, r, b]);
        if (left > l) out.push([l, top, left, bottom]);
        if (right < r) out.push([right, top, r, bottom]);
    }
    return out;
}

/** Options of the playback. */
export interface IMetafileDrawOptions {
    /** How far the whole picture reaches beyond each side of the frame, when the frame shows a crop of it. */
    crop?: { left?: number; top?: number; right?: number; bottom?: number };
    /** The CSS family list a face of the metafile is drawn with (the fonts of the page). */
    family?: (face: string) => string;
}

/**
 * Plays an EMF in a frame of the page: the picture frame of the file fills the frame, the clip of the frame cuts
 * what the records draw outside it.
 */
export function drawEmf(ctx: UniverRenderingContext, emf: IMetafile, frame: IMetafileFrame, options: IMetafileDrawOptions = {}): void {
    const crop = options.crop ?? {};
    const full = {
        left: frame.left - (crop.left ?? 0),
        top: frame.top - (crop.top ?? 0),
        width: frame.width + (crop.left ?? 0) + (crop.right ?? 0),
        height: frame.height + (crop.top ?? 0) + (crop.bottom ?? 0),
    };
    const sourceWidth = emf.frame.right - emf.frame.left;
    const sourceHeight = emf.frame.bottom - emf.frame.top;
    if (!(sourceWidth > 0 && sourceHeight > 0 && full.width > 0 && full.height > 0)) {
        return;
    }
    // device units of the reference device to the page
    const scaleX = full.width / sourceWidth;
    const scaleY = full.height / sourceHeight;
    const toPage: Matrix = [scaleX, 0, 0, scaleY, full.left - emf.frame.left * scaleX, full.top - emf.frame.top * scaleY];
    const family = options.family ?? ((face: string) => `"${face}"`);
    const whole: Rect = [emf.frame.left - 1e6, emf.frame.top - 1e6, emf.frame.right + 1e6, emf.frame.bottom + 1e6];
    const { view } = emf;
    const i32 = (offset: number) => view.getInt32(offset, true);
    const u32 = (offset: number) => view.getUint32(offset, true);
    const i16 = (offset: number) => view.getInt16(offset, true);
    const f32 = (offset: number) => view.getFloat32(offset, true);

    const initial = (): IState => ({
        mapMode: 1,
        windowOrg: [0, 0],
        windowExt: [1, 1],
        viewportOrg: [0, 0],
        viewportExt: [1, 1],
        world: [1, 0, 0, 1, 0, 0],
        pen: stockObject(7) as IPen,
        brush: stockObject(0) as IBrush,
        font: stockObject(13) as IFont,
        textColor: BLACK,
        bkColor: 'rgb(255,255,255)',
        bkMode: 2,
        textAlign: 0,
        polyFillMode: 1,
        position: [0, 0],
        clip: null,
    });
    let state = initial();
    const stack: IState[] = [];
    const objects = new Map<number, IPen | IBrush | IFont | undefined>();
    let path: Array<[string, number[]]> | null = null;

    /** Logical coordinates to device units: the world transform, then the window and the viewport. */
    const deviceMatrix = (): Matrix => {
        const isotropic = state.mapMode === 7;
        let sx = state.mapMode === 7 || state.mapMode === 8 ? state.viewportExt[0] / (state.windowExt[0] || 1) : 1;
        let sy = state.mapMode === 7 || state.mapMode === 8 ? state.viewportExt[1] / (state.windowExt[1] || 1) : 1;
        if (isotropic) {
            const s = Math.min(Math.abs(sx), Math.abs(sy));
            sx = Math.sign(sx || 1) * s;
            sy = Math.sign(sy || 1) * s;
        }
        // MM_LOMETRIC and the other metric modes: millimetres of a unit, y upwards
        const metric: Record<number, number> = { 2: 0.1, 3: 0.01, 4: 0.254, 5: 0.0254, 6: 25.4 / 1440 };
        if (metric[state.mapMode] !== undefined) {
            sx = metric[state.mapMode] * emf.perMillimeter.x;
            sy = -metric[state.mapMode] * emf.perMillimeter.y;
        }
        const page: Matrix = [sx, 0, 0, sy, state.viewportOrg[0] - state.windowOrg[0] * sx, state.viewportOrg[1] - state.windowOrg[1] * sy];
        return multiply(state.world, page);
    };
    const point = (x: number, y: number, m: Matrix): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
    const toCanvas = (m: Matrix): Matrix => multiply(m, toPage);

    const withClip = (draw: () => void) => {
        ctx.save();
        if (state.clip) {
            // the clip region is kept in device units: its rectangles on the page
            ctx.beginPath();
            for (const [l, t, r, b] of state.clip) {
                const [a, c] = point(l, t, toPage);
                const [d, e] = point(r, b, toPage);
                ctx.rect(Math.min(a, d), Math.min(c, e), Math.abs(d - a), Math.abs(e - c));
            }
            ctx.clip();
        }
        try {
            draw();
        } finally {
            ctx.restore();
        }
    };

    const strokeStyle = (pen: IPen, m: Matrix): number => {
        // a cosmetic pen is one pixel of the reference device; a geometric one its width in logical units
        const scale = Math.hypot(m[0], m[1]);
        const width = pen.cosmetic || pen.width <= 1 ? Math.max(pen.width, 1) * Math.hypot(toPage[0], toPage[1]) : pen.width * scale * Math.hypot(toPage[0], toPage[1]);
        ctx.lineWidth = width;
        ctx.strokeStyle = pen.color;
        const dashes: Record<number, number[]> = { 1: [3, 1], 2: [1, 1], 3: [3, 1, 1, 1], 4: [3, 1, 1, 1, 1, 1] };
        const dash = dashes[pen.style & 0xF];
        ctx.setLineDash(dash ? dash.map((value) => value * width * 2) : []);
        return width;
    };

    const paintPath = (trace: () => void, fill: boolean, stroke: boolean) => {
        const pen = state.pen;
        const brush = state.brush;
        const doFill = fill && brush.color !== null;
        const doStroke = stroke && pen.style !== 5;
        if (!doFill && !doStroke) {
            return;
        }
        withClip(() => {
            ctx.beginPath();
            trace();
            if (doFill) {
                ctx.fillStyle = brush.color as string;
                ctx.fill(state.polyFillMode === 2 ? 'nonzero' : 'evenodd');
            }
            if (doStroke) {
                strokeStyle(pen, deviceMatrix());
                ctx.stroke();
            }
        });
    };

    const tracePoints = (points: Array<[number, number]>, close: boolean) => {
        const m = toCanvas(deviceMatrix());
        points.forEach(([x, y], index) => {
            const [px, py] = point(x, y, m);
            if (index === 0) {
                ctx.moveTo(px, py);
            } else {
                ctx.lineTo(px, py);
            }
        });
        if (close) {
            ctx.closePath();
        }
    };

    const readPoints = (offset: number, count: number, short: boolean): Array<[number, number]> => {
        const points: Array<[number, number]> = [];
        for (let i = 0; i < count; i++) {
            points.push(short ? [i16(offset + i * 4), i16(offset + i * 4 + 2)] : [i32(offset + i * 8), i32(offset + i * 8 + 4)]);
        }
        return points;
    };

    const text = (offset: number, size: number) => {
        // EMR_EXTTEXTOUTW: bounds, graphics mode, scales, then EMRTEXT
        const at = offset + 36;
        const referenceX = i32(at);
        const referenceY = i32(at + 4);
        const count = u32(at + 8);
        const stringOffset = u32(at + 12);
        const flags = u32(at + 16);
        const rect: Rect = [i32(at + 20), i32(at + 24), i32(at + 28), i32(at + 32)];
        const dxOffset = u32(at + 36);
        if (count === 0 || offset + stringOffset + count * 2 > offset + size) {
            return;
        }
        let string = '';
        for (let i = 0; i < count; i++) {
            string += String.fromCharCode(view.getUint16(offset + stringOffset + i * 2, true));
        }
        const advances: number[] = [];
        if (dxOffset > 0 && offset + dxOffset + count * 4 <= offset + size) {
            const pairs = (flags & 0x2000) !== 0; // ETO_PDY: a pair per character
            for (let i = 0; i < count; i++) {
                advances.push(i32(offset + dxOffset + i * (pairs ? 8 : 4)));
            }
        }
        const device = deviceMatrix();
        const m = toCanvas(device);
        const font = state.font;
        // the height of the font in logical units: the em when negative, the cell (with the internal leading)
        // when positive, about 1.15 em for the fonts of Office
        const em = Math.abs(font.height || 12) / (font.height > 0 ? 1.15 : 1);
        const scaleY = Math.hypot(m[2], m[3]);
        const fontSize = Math.max(0.1, em * scaleY);
        const fontString = `${font.italic ? 'italic ' : ''}${font.weight >= 600 ? 'bold ' : ''}${fontSize}px ${family(font.face)}`;
        const updateCp = (state.textAlign & 1) === 1;
        const [x0, y0] = updateCp ? state.position : [referenceX, referenceY];
        withClip(() => {
            if ((flags & 0x4) !== 0 && rect[2] > rect[0]) {
                // ETO_CLIPPED: only inside the rectangle
                ctx.beginPath();
                const [a, b] = point(rect[0], rect[1], m);
                const [c, d] = point(rect[2], rect[3], m);
                ctx.rect(Math.min(a, c), Math.min(b, d), Math.abs(c - a), Math.abs(d - b));
                ctx.clip();
            }
            if ((flags & 0x2) !== 0 && rect[2] > rect[0]) {
                // ETO_OPAQUE: the rectangle in the background colour
                const [a, b] = point(rect[0], rect[1], m);
                const [c, d] = point(rect[2], rect[3], m);
                ctx.fillStyle = state.bkColor;
                ctx.fillRect(Math.min(a, c), Math.min(b, d), Math.abs(c - a), Math.abs(d - b));
            }
            ctx.font = fontString;
            ctx.fillStyle = state.textColor;
            const metrics = ctx.measureText('Hg');
            const ascent = (metrics as TextMetrics).fontBoundingBoxAscent ?? fontSize * 0.9;
            const descent = (metrics as TextMetrics).fontBoundingBoxDescent ?? fontSize * 0.25;
            const total = advances.length === count ? advances.reduce((sum, value) => sum + value, 0) : ctx.measureText(string).width / Math.max(1e-6, Math.hypot(m[0], m[1]));
            // the alignment of the reference point: left, centre or right; top, baseline or bottom
            const horizontal = state.textAlign & 6;
            const vertical = state.textAlign & 24;
            const startX = horizontal === 6 ? x0 - total / 2 : horizontal === 2 ? x0 - total : x0;
            let x = startX;
            const escapement = ((font.escapement || 0) / 10) * Math.PI / 180;
            // the direction the transform gives the baseline on the page: a world transform turned by Excel for the
            // titles of the axes of its charts ([0 -1 1 0]), and the turn of a font with an escapement
            const turn = Math.atan2(m[1], m[0]);
            if (escapement !== 0 || Math.abs(turn) > 1e-6) {
                // the glyphs go along the turned baseline from the reference point, each turned with it
                const [rx, ry] = point(x0, y0, m);
                const scaleX = Math.hypot(m[0], m[1]);
                const lift = vertical === 24 ? 0 : vertical === 8 ? -descent : ascent;
                ctx.save();
                ctx.translate(rx, ry);
                ctx.rotate(turn - escapement);
                for (let i = 0; i < count; i++) {
                    ctx.fillText(string[i], (x - x0) * scaleX, lift);
                    x += advances.length === count ? advances[i] : ctx.measureText(string[i]).width / Math.max(1e-6, scaleX);
                }
                ctx.restore();
                return;
            }
            for (let i = 0; i < count; i++) {
                const [px, py] = point(x, y0, m);
                const baseline = vertical === 24 ? py : vertical === 8 ? py - descent : py + ascent;
                ctx.fillText(string[i], px, baseline);
                x += advances.length === count ? advances[i] : ctx.measureText(string[i]).width / Math.max(1e-6, Math.hypot(m[0], m[1]));
            }
            if (font.underline || font.strikeOut) {
                const [ax, ay] = point(startX, y0, m);
                const [bx] = point(startX + total, y0, m);
                const baseline = vertical === 24 ? ay : vertical === 8 ? ay - descent : ay + ascent;
                ctx.fillStyle = state.textColor;
                const thickness = Math.max(0.5, fontSize / 18);
                if (font.underline) {
                    ctx.fillRect(ax, baseline + fontSize * 0.1, bx - ax, thickness);
                }
                if (font.strikeOut) {
                    ctx.fillRect(ax, baseline - fontSize * 0.28, bx - ax, thickness);
                }
            }
            if (updateCp) {
                state.position = [startX + total, y0];
            }
        });
    };

    /** The bits of a monochrome pattern brush, with its two colours (null: the text and background colours). */
    const brushPattern = (offset: number, mono: boolean): IBrushPattern | undefined => {
        const bmi = offset + u32(offset + 16);
        const bits = offset + u32(offset + 24);
        if (bmi + 40 > view.byteLength) return undefined;
        const width = i32(bmi + 4);
        const height = i32(bmi + 8);
        const rows = Math.abs(height);
        const bitCount = view.getUint16(bmi + 14, true);
        const compression = u32(bmi + 16);
        if (width <= 0 || rows === 0 || width > 64 || rows > 64) {
            return undefined;
        }
        if (bitCount === 24 || bitCount === 32) {
            // a pattern in colour (Excel writes its dotted borders with 8 x 8 pixels of 32 bits and the masks of their
            // channels): the colour of each pixel
            const fields = compression === 3;
            if (compression !== 0 && !(fields && bitCount === 32)) return undefined;
            const masks = fields ? [u32(bmi + u32(bmi)), u32(bmi + u32(bmi) + 4), u32(bmi + u32(bmi) + 8)] : [];
            const channel = (value: number, mask: number) => {
                if (!mask) return 0;
                let shift = 0;
                while (((mask >>> shift) & 1) === 0 && shift < 32) shift++;
                return Math.round(((value & mask) >>> shift) * 255 / (mask >>> shift));
            };
            const stride24 = Math.ceil((width * bitCount) / 32) * 4;
            if (bits + stride24 * rows > view.byteLength) return undefined;
            const pixels: string[] = [];
            for (let row = 0; row < rows; row++) {
                const source = bits + (height > 0 ? rows - 1 - row : row) * stride24;
                for (let column = 0; column < width; column++) {
                    const at = source + column * (bitCount / 8);
                    const rgb = fields
                        ? [channel(u32(at), masks[0]), channel(u32(at), masks[1]), channel(u32(at), masks[2])]
                        : [view.getUint8(at + 2), view.getUint8(at + 1), view.getUint8(at)];
                    pixels.push(`rgb(${rgb[0]},${rgb[1]},${rgb[2]})`);
                }
            }
            return { width, height: rows, bits: new Uint8Array(0), colors: null, pixels };
        }
        if (bitCount !== 1 || compression !== 0) {
            return undefined;
        }
        const stride = Math.ceil(width / 32) * 4;
        if (bits + stride * rows > view.byteLength) return undefined;
        const out = new Uint8Array(width * rows);
        for (let row = 0; row < rows; row++) {
            // rows go from the bottom unless the height is negative
            const source = bits + (height > 0 ? rows - 1 - row : row) * stride;
            for (let column = 0; column < width; column++) {
                out[row * width + column] = (view.getUint8(source + (column >> 3)) >> (7 - (column & 7))) & 1;
            }
        }
        const palette = bmi + u32(bmi);
        const rgb = (at: number) => `rgb(${view.getUint8(at + 2)},${view.getUint8(at + 1)},${view.getUint8(at)})`;
        const colors: [string, string] | null = !mono && u32(offset + 12) === 0 && palette + 8 <= view.byteLength
            ? [rgb(palette), rgb(palette + 4)]
            : null;
        return { width, height: rows, bits: out, colors };
    };

    /**
     * A rectangle filled with a pattern brush, pixel by pixel of the device as GDI copies the pattern from the
     * origin of the device: the dotted borders of Excel. A large rectangle takes the grey of the brush.
     */
    const fillPattern = (pattern: IBrushPattern, x: number, y: number, cx: number, cy: number) => {
        const device = deviceMatrix();
        const [a, b] = point(x, y, device);
        const [c, d] = point(x + cx, y + cy, device);
        const left = Math.round(Math.min(a, c));
        const right = Math.round(Math.max(a, c));
        const top = Math.round(Math.min(b, d));
        const bottom = Math.round(Math.max(b, d));
        if ((right - left) * (bottom - top) > 40000) {
            return false;
        }
        const colors = pattern.colors ?? [state.textColor, state.bkColor];
        withClip(() => {
            for (let row = top; row < bottom; row++) {
                const line = ((row % pattern.height) + pattern.height) % pattern.height;
                let start = left;
                // the colour of a pixel of the device: its pixel of the pattern, from the origin of the device
                const colorAt = (value: number) => {
                    const index = line * pattern.width + (((value % pattern.width) + pattern.width) % pattern.width);
                    return pattern.pixels ? pattern.pixels[index] : colors[pattern.bits[index]];
                };
                for (let column = left; column <= right; column++) {
                    if (column < right && colorAt(column) === colorAt(start)) continue;
                    const [px, py] = point(start, row, toPage);
                    const [qx, qy] = point(column, row + 1, toPage);
                    ctx.fillStyle = colorAt(start);
                    ctx.fillRect(Math.min(px, qx), Math.min(py, qy), Math.abs(qx - px), Math.abs(qy - py));
                    start = column;
                }
            }
        });
        return true;
    };

    const blit = (offset: number, stretch: boolean) => {
        // EMR_BITBLT and EMR_STRETCHBLT: a pattern copy without a source fills the destination with the brush
        const x = i32(offset + 24);
        const y = i32(offset + 28);
        const cx = i32(offset + 32);
        const cy = i32(offset + 36);
        const rop = u32(offset + 40);
        // the source bitmap: offBmiSrc, cbBmiSrc, offBitsSrc (the same places in both records)
        if (u32(offset + 88) !== 0) {
            bitmap(offset + u32(offset + 84), offset + u32(offset + 92), [x, y, x + cx, y + cy]);
            return;
        }
        void stretch;
        if (rop === 0x00F00021 && state.brush.pattern && fillPattern(state.brush.pattern, x, y, cx, cy)) {
            return;
        }
        const fill = rop === 0x00F00021 ? state.brush.color : rop === 0x00000042 ? BLACK : rop === 0x00FF0062 ? 'rgb(255,255,255)' : null;
        if (fill === null) {
            return;
        }
        withClip(() => {
            const m = toCanvas(deviceMatrix());
            const [a, b] = point(x, y, m);
            const [c, d] = point(x + cx, y + cy, m);
            ctx.fillStyle = fill;
            ctx.fillRect(Math.min(a, c), Math.min(b, d), Math.abs(c - a), Math.abs(d - b));
        });
    };

    const bitmap = (bmi: number, bits: number, dest: Rect) => {
        const width = i32(bmi + 4);
        const height = i32(bmi + 8);
        const bitCount = view.getUint16(bmi + 14, true);
        const compression = u32(bmi + 16);
        const headerSize = u32(bmi);
        // BI_RGB, and BI_BITFIELDS for 16 and 32 bits (the masks of each channel after the header, as GDI+ writes)
        const fields = compression === 3 && (bitCount === 16 || bitCount === 32);
        if (width <= 0 || height === 0 || (compression !== 0 && !fields) || ![1, 4, 8, 16, 24, 32].includes(bitCount) ||
            (bitCount === 16 && !fields) || typeof document === 'undefined') {
            return;
        }
        const masks = fields ? [u32(bmi + headerSize), u32(bmi + headerSize + 4), u32(bmi + headerSize + 8)] : [];
        const channel = (value: number, mask: number) => {
            if (!mask) return 0;
            let shift = 0;
            while (((mask >>> shift) & 1) === 0 && shift < 32) shift++;
            const top = mask >>> shift;
            return Math.round(((value & mask) >>> shift) * 255 / top);
        };
        const rows = Math.abs(height);
        const stride = Math.ceil((width * bitCount) / 32) * 4;
        if (bits + stride * rows > view.byteLength) {
            return;
        }
        const colors = bitCount <= 8 ? (u32(bmi + 32) || 1 << bitCount) : 0;
        const palette = (index: number): [number, number, number] => {
            const at = bmi + headerSize + index * 4;
            return at + 3 <= view.byteLength ? [view.getUint8(at + 2), view.getUint8(at + 1), view.getUint8(at)] : [0, 0, 0];
        };
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = rows;
        const context = canvas.getContext('2d');
        if (!context) {
            return;
        }
        const image = context.createImageData(width, rows);
        for (let row = 0; row < rows; row++) {
            const source = bits + (height > 0 ? rows - 1 - row : row) * stride;
            for (let column = 0; column < width; column++) {
                let rgb: [number, number, number];
                if (fields) {
                    const value = bitCount === 32 ? u32(source + column * 4) : view.getUint16(source + column * 2, true);
                    rgb = [channel(value, masks[0]), channel(value, masks[1]), channel(value, masks[2])];
                } else if (bitCount === 24 || bitCount === 32) {
                    const at = source + column * (bitCount / 8);
                    rgb = [view.getUint8(at + 2), view.getUint8(at + 1), view.getUint8(at)];
                } else {
                    const bit = column * bitCount;
                    const byte = view.getUint8(source + (bit >> 3));
                    const index = (byte >> (8 - bitCount - (bit & 7))) & ((1 << bitCount) - 1);
                    rgb = index < colors ? palette(index) : [0, 0, 0];
                }
                const target = (row * width + column) * 4;
                image.data[target] = rgb[0];
                image.data[target + 1] = rgb[1];
                image.data[target + 2] = rgb[2];
                image.data[target + 3] = 255;
            }
        }
        context.putImageData(image, 0, 0);
        withClip(() => {
            const m = toCanvas(deviceMatrix());
            const [a, b] = point(dest[0], dest[1], m);
            const [c, d] = point(dest[2], dest[3], m);
            ctx.drawImage(canvas, Math.min(a, c), Math.min(b, d), Math.abs(c - a), Math.abs(d - b));
        });
    };

    ctx.save();
    // what the records draw outside the frame of the picture is not shown
    ctx.beginPath();
    ctx.rect(frame.left, frame.top, frame.width, frame.height);
    ctx.clip();
    try {
        for (const { type, offset, size } of emf.records) {
            switch (type) {
                case 9:
                    state.windowExt = [i32(offset + 8), i32(offset + 12)];
                    break;
                case 10:
                    state.windowOrg = [i32(offset + 8), i32(offset + 12)];
                    break;
                case 11:
                    state.viewportExt = [i32(offset + 8), i32(offset + 12)];
                    break;
                case 12:
                    state.viewportOrg = [i32(offset + 8), i32(offset + 12)];
                    break;
                case 17:
                    state.mapMode = u32(offset + 8);
                    break;
                case 18:
                    state.bkMode = u32(offset + 8);
                    break;
                case 19:
                    state.polyFillMode = u32(offset + 8);
                    break;
                case 22:
                    state.textAlign = u32(offset + 8);
                    break;
                case 24:
                    state.textColor = colorRef(u32(offset + 8));
                    break;
                case 25:
                    state.bkColor = colorRef(u32(offset + 8));
                    break;
                case 27:
                    state.position = [i32(offset + 8), i32(offset + 12)];
                    // in a path, a move starts a new figure there
                    path?.push(['M', state.position]);
                    break;
                case 33:
                    stack.push({ ...state, position: [...state.position], clip: state.clip ? state.clip.map((rect) => [...rect] as Rect) : null });
                    break;
                case 34: {
                    const back = i32(offset + 8);
                    const index = back < 0 ? stack.length + back : back;
                    if (index >= 0 && index < stack.length) {
                        state = stack[index];
                        stack.length = index;
                    }
                    break;
                }
                case 35:
                    state.world = [f32(offset + 8), f32(offset + 12), f32(offset + 16), f32(offset + 20), f32(offset + 24), f32(offset + 28)];
                    break;
                case 36: {
                    const m: Matrix = [f32(offset + 8), f32(offset + 12), f32(offset + 16), f32(offset + 20), f32(offset + 24), f32(offset + 28)];
                    const mode = u32(offset + 32);
                    state.world = mode === 1 ? [1, 0, 0, 1, 0, 0] : mode === 2 ? multiply(m, state.world) : mode === 3 ? multiply(state.world, m) : mode === 4 ? m : state.world;
                    break;
                }
                case 37: {
                    const handle = u32(offset + 8);
                    const object = (handle & 0x80000000) !== 0 ? stockObject(handle & 0x7FFFFFFF) : objects.get(handle);
                    if (object && 'cosmetic' in object) state.pen = object as IPen;
                    else if (object && 'face' in object) state.font = object as IFont;
                    else if (object) state.brush = object as IBrush;
                    break;
                }
                case 38:
                    objects.set(u32(offset + 8), { style: u32(offset + 12), width: i32(offset + 16), color: colorRef(u32(offset + 24)), cosmetic: i32(offset + 16) <= 1 });
                    break;
                case 95: {
                    const style = u32(offset + 28);
                    const width = u32(offset + 32);
                    objects.set(u32(offset + 8), { style: style & 0xF, width, color: colorRef(u32(offset + 40)), cosmetic: (style & 0x00010000) === 0 && width <= 1 });
                    break;
                }
                case 39: {
                    const style = u32(offset + 12);
                    objects.set(u32(offset + 8), { color: style === 1 ? null : colorRef(u32(offset + 16)) });
                    break;
                }
                case 93:
                case 94:
                    // monochrome and DIB pattern brushes: the dots of Excel's dotted borders; grey where the pattern
                    // cannot be read
                    objects.set(u32(offset + 8), { color: 'rgb(128,128,128)', pattern: brushPattern(offset, type === 93) });
                    break;
                case 82: {
                    const at = offset + 12;
                    let face = '';
                    for (let i = 0; i < 32; i++) {
                        const code = view.getUint16(at + 28 + i * 2, true);
                        if (code === 0) break;
                        face += String.fromCharCode(code);
                    }
                    objects.set(u32(offset + 8), {
                        height: i32(at),
                        escapement: i32(at + 8),
                        weight: i32(at + 16),
                        italic: view.getUint8(at + 20) !== 0,
                        underline: view.getUint8(at + 21) !== 0,
                        strikeOut: view.getUint8(at + 22) !== 0,
                        face: face || 'Arial',
                    });
                    break;
                }
                case 40:
                    objects.delete(u32(offset + 8));
                    break;
                case 30: case 29: {
                    const m = deviceMatrix();
                    const [a, b] = point(i32(offset + 8), i32(offset + 12), m);
                    const [c, d] = point(i32(offset + 16), i32(offset + 20), m);
                    const rect: Rect = [Math.min(a, c), Math.min(b, d), Math.max(a, c), Math.max(b, d)];
                    state.clip = type === 30 ? intersect(state.clip, rect) : subtract(state.clip, rect, whole);
                    break;
                }
                case 75: {
                    const bytes = u32(offset + 8);
                    const mode = u32(offset + 12);
                    if (bytes === 0) {
                        if (mode === 5) state.clip = null;
                        break;
                    }
                    // RGNDATA: header (32 bytes), then rectangles in device units
                    const header = offset + 16;
                    const count = u32(header + 8);
                    const rects: Rect[] = [];
                    for (let i = 0; i < count; i++) {
                        const at = header + 32 + i * 16;
                        rects.push([i32(at), i32(at + 4), i32(at + 8), i32(at + 12)]);
                    }
                    if (mode === 5 || (state.clip === null && mode === 1)) state.clip = rects;
                    else if (mode === 1) state.clip = rects.flatMap((rect) => intersect(state.clip, rect));
                    else if (mode === 2) state.clip = [...(state.clip ?? []), ...rects];
                    else if (mode === 4) state.clip = rects.reduce<Rect[] | null>((region, rect) => subtract(region, rect, whole), state.clip);
                    break;
                }
                case 54: {
                    const from = state.position;
                    const to: [number, number] = [i32(offset + 8), i32(offset + 12)];
                    if (path) {
                        path.push(['L', to]);
                    } else {
                        paintPath(() => tracePoints([from, to], false), false, true);
                    }
                    state.position = to;
                    break;
                }
                case 43: case 42: {
                    const r: Rect = [i32(offset + 8), i32(offset + 12), i32(offset + 16), i32(offset + 20)];
                    paintPath(() => {
                        if (type === 43) {
                            tracePoints([[r[0], r[1]], [r[2], r[1]], [r[2], r[3]], [r[0], r[3]]], true);
                            return;
                        }
                        const m = toCanvas(deviceMatrix());
                        const [a, b] = point(r[0], r[1], m);
                        const [c, d] = point(r[2], r[3], m);
                        ctx.ellipse((a + c) / 2, (b + d) / 2, Math.abs(c - a) / 2, Math.abs(d - b) / 2, 0, 0, Math.PI * 2);
                    }, true, true);
                    break;
                }
                case 3: case 4: case 86: case 87: case 6: case 89: {
                    const short = type >= 85;
                    const count = u32(offset + 24);
                    const points = readPoints(offset + 28, count, short);
                    const polygon = type === 3 || type === 86;
                    const continues = type === 6 || type === 89;
                    if (path) {
                        points.forEach((p, index) => path!.push([index === 0 && !continues ? 'M' : 'L', p]));
                        if (polygon) path.push(['Z', []]);
                    } else {
                        const all = continues ? [state.position, ...points] : points;
                        paintPath(() => tracePoints(all, polygon), polygon, true);
                    }
                    if (count > 0) state.position = points[count - 1];
                    break;
                }
                case 8: case 7: case 91: case 90: {
                    const short = type >= 85;
                    const polygons = u32(offset + 24);
                    const total = u32(offset + 28);
                    const counts: number[] = [];
                    for (let i = 0; i < polygons; i++) counts.push(u32(offset + 32 + i * 4));
                    const points = readPoints(offset + 32 + polygons * 4, total, short);
                    const fill = type === 8 || type === 91;
                    paintPath(() => {
                        let at = 0;
                        for (const n of counts) {
                            tracePoints(points.slice(at, at + n), fill);
                            at += n;
                        }
                    }, fill, true);
                    break;
                }
                case 5: case 88: case 2: case 85: {
                    const short = type >= 85;
                    const count = u32(offset + 24);
                    const points = readPoints(offset + 28, count, short);
                    const to = type === 5 || type === 88;
                    const all = to ? [state.position, ...points] : points;
                    // in a path, POLYBEZIERTO goes on from the current point of the figure, POLYBEZIER starts one
                    const trace = () => {
                        if (!all.length) return;
                        const m = toCanvas(deviceMatrix());
                        const [sx, sy] = point(all[0][0], all[0][1], m);
                        ctx.moveTo(sx, sy);
                        for (let i = 1; i + 2 < all.length; i += 3) {
                            const [ax, ay] = point(all[i][0], all[i][1], m);
                            const [bx, by] = point(all[i + 1][0], all[i + 1][1], m);
                            const [cx, cy] = point(all[i + 2][0], all[i + 2][1], m);
                            ctx.bezierCurveTo(ax, ay, bx, by, cx, cy);
                        }
                    };
                    if (path) {
                        path.push(to ? ['C', points.flat()] : ['B', points.flat()]);
                    } else {
                        paintPath(trace, false, true);
                    }
                    if (count > 0) state.position = points[count - 1];
                    break;
                }
                case 59:
                    path = [];
                    break;
                case 68:
                    // EMR_ABORTPATH: the path is dropped
                    path = null;
                    break;
                case 61:
                    path?.push(['Z', []]);
                    break;
                case 60:
                    break;
                case 62: case 63: case 64: {
                    const recorded = path ?? [];
                    path = null;
                    paintPath(() => {
                        const m = toCanvas(deviceMatrix());
                        for (const [op, values] of recorded) {
                            if (op === 'Z') {
                                ctx.closePath();
                            } else if (op === 'B' || op === 'C') {
                                // B starts a figure at its first point, C goes on from the current point
                                const pts: Array<[number, number]> = [];
                                for (let i = 0; i + 1 < values.length; i += 2) pts.push(point(values[i], values[i + 1], m));
                                if (!pts.length) continue;
                                const first = op === 'B' ? 1 : 0;
                                if (op === 'B') ctx.moveTo(pts[0][0], pts[0][1]);
                                for (let i = first; i + 2 < pts.length; i += 3) {
                                    ctx.bezierCurveTo(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], pts[i + 2][0], pts[i + 2][1]);
                                }
                            } else {
                                const [px, py] = point(values[0], values[1], m);
                                if (op === 'M') ctx.moveTo(px, py);
                                else ctx.lineTo(px, py);
                            }
                        }
                    }, type !== 64, type !== 62);
                    break;
                }
                case 84:
                    text(offset, size);
                    break;
                case 76:
                    blit(offset, false);
                    break;
                case 77:
                    blit(offset, true);
                    break;
                case 81: {
                    // EMR_STRETCHDIBITS
                    const x = i32(offset + 24);
                    const y = i32(offset + 28);
                    const bmiOffset = u32(offset + 48);
                    const bitsOffset = u32(offset + 56);
                    const cx = i32(offset + 72);
                    const cy = i32(offset + 76);
                    if (u32(offset + 52) > 0) bitmap(offset + bmiOffset, offset + bitsOffset, [x, y, x + cx, y + cy]);
                    break;
                }
                default:
                    break;
            }
        }
    } finally {
        ctx.restore();
    }
}
