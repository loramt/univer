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

import type { IDocDrawingBase, IDocumentBody, IDocumentData, IDocumentStyle, Nullable } from '@univerjs/core';
import {
    AlignTypeH,
    AlignTypeV,
    DocumentDataModel,
    ObjectRelativeFromH,
    ObjectRelativeFromV,
    PositionedObjectLayoutType,
} from '@univerjs/core';
import type { IDocumentSkeletonDrawing, IDocumentSkeletonPage } from '../../basics/i-document-skeleton-cached';
import type { UniverRenderingContext } from '../../context';
import type { IDocumentsConfig } from './doc-component';
import { DocumentSkeleton } from './layout/doc-skeleton';
import { DocumentViewModel } from './view-model/document-view-model';

/** The outline and the fill of a text box, as Word draws its shape: the line centred on the frame. */
export interface IDocTextBoxShape {
    /** A preset geometry of Word: rect, roundRect, flowChartAlternateProcess, ellipse (others are drawn as rect). */
    geometry?: string;
    /** The corner radius of roundRect and flowChartAlternateProcess, as a fraction of the shortest side. */
    adjust?: number;
    fill?: Nullable<{ color: string; opacity?: number }>;
    line?: Nullable<{ color: string; opacity?: number; width: number; dash?: string }>;
}

/**
 * A drawing with a content of its own (a text box of Word): the content is laid out in the frame of the drawing,
 * within its insets, and anchored to its top, middle or bottom; the drawing keeps its place in the flow as any other.
 */
export interface IDocTextBox {
    body: IDocumentBody;
    /** From the edges of the frame to the content (px): Word's insets with the text rectangle of the shape. */
    insets?: { left?: number; top?: number; right?: number; bottom?: number };
    verticalAlign?: 'top' | 'middle' | 'bottom';
    /** False when the lines of the content do not wrap (bodyPr wrap="none"). */
    wrap?: boolean;
    shape?: IDocTextBoxShape;
}

export interface IDocTextBoxFrame {
    left: number;
    top: number;
    width: number;
    height: number;
}

/** The text box of a drawing, when it has one. */
export function getDocTextBox(drawing: Nullable<IDocDrawingBase>): IDocTextBox | undefined {
    const textBox = (drawing as Nullable<IDocDrawingBase & { textBox?: unknown }>)?.textBox as Nullable<IDocTextBox>;
    return textBox && typeof textBox === 'object' && typeof textBox.body?.dataStream === 'string' ? textBox : undefined;
}

/**
 * Where the frame of a drawing is on its page, from the top-left corner of the page, as the drawings of the
 * document are placed: from the page for offsets and alignments relative to the page or its margins, else from
 * the origin of the segment the drawing is laid out in (the body below the top margin, a header, a footer).
 */
export function getDocTextBoxFrame(
    drawing: IDocumentSkeletonDrawing,
    segmentTop: number,
    segmentLeft: number,
    anchorPage?: Pick<IDocumentSkeletonPage, 'pageWidth' | 'pageHeight' | 'marginLeft' | 'marginRight'>
): IDocTextBoxFrame {
    const origin = drawing.drawingOrigin;
    const width = origin.docTransform?.size?.width ?? drawing.width;
    const height = origin.docTransform?.size?.height ?? drawing.height;
    let left = drawing.aLeft + segmentLeft;
    let top = drawing.aTop + segmentTop;
    if (anchorPage && origin.layoutType !== PositionedObjectLayoutType.INLINE) {
        const positionH = origin.docTransform?.positionH;
        if (positionH && (positionH.relativeFrom === ObjectRelativeFromH.PAGE || positionH.relativeFrom === ObjectRelativeFromH.MARGIN)) {
            const areaLeft = positionH.relativeFrom === ObjectRelativeFromH.MARGIN ? anchorPage.marginLeft ?? 0 : 0;
            const areaWidth = positionH.relativeFrom === ObjectRelativeFromH.MARGIN
                ? anchorPage.pageWidth - areaLeft - (anchorPage.marginRight ?? 0)
                : anchorPage.pageWidth;
            left = positionH.align === AlignTypeH.RIGHT
                ? areaLeft + areaWidth - width
                : positionH.align === AlignTypeH.CENTER
                    ? areaLeft + areaWidth / 2 - width / 2
                    : positionH.align != null
                        ? areaLeft
                        : areaLeft + (positionH.posOffset ?? (positionH.percent ?? 0) * areaWidth);
        }
        const positionV = origin.docTransform?.positionV;
        if (positionV && positionV.relativeFrom === ObjectRelativeFromV.PAGE && Number.isFinite(anchorPage.pageHeight)) {
            top = positionV.align === AlignTypeV.BOTTOM
                ? anchorPage.pageHeight - height
                : positionV.align === AlignTypeV.CENTER
                    ? anchorPage.pageHeight / 2 - height / 2
                    : positionV.align != null
                        ? 0
                        : positionV.posOffset ?? (positionV.percent ?? 0) * anchorPage.pageHeight;
        }
    }
    return { left, top, width, height };
}

/** A CSS colour with an opacity, for the canvas. */
function paint(color: string, opacity = 1): string {
    const hex = /^#?([0-9a-f]{6})$/i.exec(color.trim())?.[1];
    if (!hex) {
        return color;
    }
    const value = Number.parseInt(hex, 16);
    return `rgba(${(value >> 16) & 255}, ${(value >> 8) & 255}, ${value & 255}, ${Math.max(0, Math.min(1, opacity))})`;
}

/** The dashes of Word's preset line dashes, in line widths (ECMA-376 prstDash, as Word draws them). */
const DASHES: Record<string, number[]> = {
    dot: [1, 1],
    sysDot: [1, 1],
    dash: [4, 3],
    sysDash: [3, 1],
    lgDash: [8, 3],
    dashDot: [4, 3, 1, 3],
    sysDashDot: [3, 1, 1, 1],
    lgDashDot: [8, 3, 1, 3],
    lgDashDotDot: [8, 3, 1, 3, 1, 3],
    sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** The outline of a shape in its frame, as a path of the context. */
function traceShape(ctx: UniverRenderingContext, shape: IDocTextBoxShape, frame: IDocTextBoxFrame): void {
    const { left, top, width, height } = frame;
    const geometry = shape.geometry ?? 'rect';
    if (geometry === 'ellipse') {
        // four cubic arcs: the circle of Bézier, k = 4(√2 - 1)/3
        const k = 0.5522847498;
        const rx = width / 2;
        const ry = height / 2;
        const cx = left + rx;
        const cy = top + ry;
        ctx.moveTo(cx + rx, cy);
        ctx.bezierCurveTo(cx + rx, cy + ry * k, cx + rx * k, cy + ry, cx, cy + ry);
        ctx.bezierCurveTo(cx - rx * k, cy + ry, cx - rx, cy + ry * k, cx - rx, cy);
        ctx.bezierCurveTo(cx - rx, cy - ry * k, cx - rx * k, cy - ry, cx, cy - ry);
        ctx.bezierCurveTo(cx + rx * k, cy - ry, cx + rx, cy - ry * k, cx + rx, cy);
        ctx.closePath();
        return;
    }
    const rounded = geometry === 'roundRect' || geometry === 'flowChartAlternateProcess';
    const radius = rounded ? Math.min(width, height) * Math.max(0, Math.min(0.5, shape.adjust ?? 1 / 6)) : 0;
    if (radius <= 0) {
        ctx.rect(left, top, width, height);
        return;
    }
    // quarter circles as cubic arcs, k = 4(√2 - 1)/3
    const c = radius * (1 - 0.5522847498);
    const right = left + width;
    const bottom = top + height;
    ctx.moveTo(left + radius, top);
    ctx.lineTo(right - radius, top);
    ctx.bezierCurveTo(right - c, top, right, top + c, right, top + radius);
    ctx.lineTo(right, bottom - radius);
    ctx.bezierCurveTo(right, bottom - c, right - c, bottom, right - radius, bottom);
    ctx.lineTo(left + radius, bottom);
    ctx.bezierCurveTo(left + c, bottom, left, bottom - c, left, bottom - radius);
    ctx.lineTo(left, top + radius);
    ctx.bezierCurveTo(left, top + c, left + c, top, left + radius, top);
    ctx.closePath();
}

/** Fills and strokes the shape of a text box: Word centres its line on the frame. */
export function drawDocTextBoxShape(ctx: UniverRenderingContext, shape: Nullable<IDocTextBoxShape>, frame: IDocTextBoxFrame): void {
    const fill = shape?.fill;
    const line = shape?.line;
    if (!shape || (!fill && !(line && line.width > 0)) || frame.width <= 0 || frame.height <= 0) {
        return;
    }
    ctx.save();
    ctx.beginPath();
    traceShape(ctx, shape, frame);
    if (fill) {
        ctx.fillStyle = paint(fill.color, fill.opacity);
        ctx.fill();
    }
    if (line && line.width > 0) {
        ctx.lineWidth = line.width;
        ctx.strokeStyle = paint(line.color, line.opacity);
        const dashes = DASHES[line.dash ?? ''];
        ctx.setLineDash(dashes ? dashes.map((dash) => dash * line.width) : []);
        ctx.stroke();
    }
    ctx.restore();
}

/** The height of the page a text box is laid out on: longer than any content, so that it has one page. */
const TEXT_BOX_PAGE_HEIGHT = 100000;

interface ITextBoxEntry {
    body: IDocumentBody;
    key: string;
    skeleton: DocumentSkeleton;
    documents: ITextBoxDocuments;
    contentHeight: number;
}

/** What the painter needs of the documents that draws a text box (Documents, made by the host painter). */
export interface ITextBoxDocuments {
    draw(ctx: UniverRenderingContext): void;
    setPageFieldContext(context: IDocumentsConfig['pageFieldContext']): void;
    dispose(): void;
}

/**
 * Lays out and draws the content of the text boxes of a document: each one is a document of its own, one page as
 * as wide as the frame with the insets as margins, laid out with the settings of the host document and drawn with
 * its painter where the frame is. What does not fit in the room inside the insets is cut there, as in Word.
 */
export class DocTextBoxPainter {
    private readonly _entries = new Map<string, ITextBoxEntry>();

    constructor(private readonly _createDocuments: (key: string, skeleton: DocumentSkeleton) => ITextBoxDocuments) {}

    draw(
        ctx: UniverRenderingContext,
        drawingId: string,
        textBox: IDocTextBox,
        frame: IDocTextBoxFrame,
        host: DocumentSkeleton,
        pageFieldContext?: IDocumentsConfig['pageFieldContext']
    ): void {
        if (frame.width <= 0 || frame.height <= 0) {
            return;
        }
        const entry = this._entry(drawingId, textBox, frame, host);
        if (!entry) {
            return;
        }
        const insets = textBox.insets ?? {};
        const top = insets.top ?? 0;
        const room = frame.height - top - (insets.bottom ?? 0);
        const share = textBox.verticalAlign === 'middle' ? 0.5 : textBox.verticalAlign === 'bottom' ? 1 : 0;
        // The page of the content starts with the top inset: middle and bottom move it by what is left of the room.
        const offset = share > 0 ? (room - entry.contentHeight) * share : 0;
        entry.documents.setPageFieldContext(pageFieldContext);
        ctx.save();
        ctx.beginPath();
        // Word cuts the lines that do not fit at the bottom of the room for the text, inside the insets
        ctx.rect(frame.left, frame.top + top, frame.width, Math.max(0, room));
        ctx.clip();
        ctx.translate(frame.left, frame.top + offset);
        try {
            entry.documents.draw(ctx);
        } finally {
            ctx.restore();
        }
    }

    dispose(): void {
        for (const entry of this._entries.values()) {
            entry.documents.dispose();
            entry.skeleton.dispose();
        }
        this._entries.clear();
    }

    private _entry(drawingId: string, textBox: IDocTextBox, frame: IDocTextBoxFrame, host: DocumentSkeleton): ITextBoxEntry | undefined {
        const insets = textBox.insets ?? {};
        const key = JSON.stringify([frame.width, frame.height, insets, textBox.wrap !== false]);
        const cached = this._entries.get(drawingId);
        if (cached && cached.key === key && cached.body === textBox.body) {
            return cached;
        }
        if (cached) {
            cached.documents.dispose();
            cached.skeleton.dispose();
            this._entries.delete(drawingId);
        }
        const snapshot = host.getViewModel().getSnapshot();
        const {
            defaultHeaderId: _defaultHeader,
            defaultFooterId: _defaultFooter,
            evenPageHeaderId: _evenHeader,
            evenPageFooterId: _evenFooter,
            firstPageHeaderId: _firstHeader,
            firstPageFooterId: _firstFooter,
            renderConfig: _renderConfig,
            ...style
        } = snapshot.documentStyle ?? ({} as IDocumentStyle);
        const data: IDocumentData = {
            ...snapshot,
            id: `${snapshot.id}::text-box::${drawingId}`,
            body: textBox.body,
            documentStyle: {
                ...style,
                // One page as long as the content: Word draws the lines that do not fit cut by the frame. A content
                // that does not wrap is laid out on lines as long as it takes.
                pageSize: { width: textBox.wrap === false ? Math.max(frame.width, 100000) : frame.width, height: TEXT_BOX_PAGE_HEIGHT },
                marginTop: insets.top ?? 0,
                marginBottom: insets.bottom ?? 0,
                marginLeft: insets.left ?? 0,
                marginRight: insets.right ?? 0,
                marginHeader: 0,
                marginFooter: 0,
                useFirstPageHeaderFooter: 0,
                evenAndOddHeaders: 0,
                pageNumberStart: 1,
            } as IDocumentStyle,
            headers: {},
            footers: {},
            drawings: {},
            drawingsOrder: [],
        };
        try {
            const skeleton = DocumentSkeleton.create(new DocumentViewModel(new DocumentDataModel(data)), host.getLocaleService());
            skeleton.calculate();
            const page = skeleton.getSkeletonData()?.pages[0];
            const entry: ITextBoxEntry = {
                body: textBox.body,
                key,
                skeleton,
                documents: this._createDocuments(data.id, skeleton),
                contentHeight: page?.height ?? 0,
            };
            this._entries.set(drawingId, entry);
            return entry;
        } catch (error) {
            console.warn(`The text box ${drawingId} could not be laid out.`, error);
            return undefined;
        }
    }
}
