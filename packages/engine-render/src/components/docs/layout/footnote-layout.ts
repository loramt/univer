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

import type { IDocumentNoteDecoration } from '@univerjs/core';
import { ColumnSeparatorType, CustomRangeType, DataStreamTreeTokenType, DocumentDataModel, GridType } from '@univerjs/core';
import type {
    IDocumentSkeletonGlyph,
    IDocumentSkeletonNote,
    IDocumentSkeletonNoteDecoration,
    IDocumentSkeletonPage,
    IDocumentSkeletonTable,
} from '../../../basics/i-document-skeleton-cached';
import type { IParagraphConfig, ISectionBreakConfig } from '../../../basics/interfaces';
import type { INoteReferenceLayout } from './note-numbering';
import type { ILayoutContext } from './tools';
import { DocumentSkeletonPageType } from '../../../basics/i-document-skeleton-cached';
import { getDocumentCompatibilityPolicy } from '../document-compatibility';
import { DocumentViewModel } from '../view-model/document-view-model';
import { dealWithSection } from './block/section';
import { createSkeletonLetterGlyph } from './model/glyph';
import {
    createSkeletonPage,
    expandCellPageHeightForFlowTables,
    expandCellPageHeightForInlineDrawings,
    getNextSkeletonPageNumber,
} from './model/page';
import { formatNoteNumber } from './note-numbering';
import {
    clearFontCreateConfigCache,
    getFontConfigFromLastGlyph,
    getNullSkeleton,
    updateBlockIndex,
    updateInlineDrawingCoordsAndBorder,
} from './tools';

export interface INoteLayoutConstraints {
    width: number;
    firstPageHeight: number;
    continuationPageHeight: number;
    columnCount?: number;
    firstColumn?: { index: number; top: number };
}

function createFootnoteColumns(config: ISectionBreakConfig, width: number, count: number) {
    if (count <= 1) {
        return [];
    }
    if (config.columnProperties?.length === count) {
        return config.columnProperties;
    }
    const gap = Math.min(48, width / (count * 2));
    const columnWidth = Math.max(1, (width - gap * (count - 1)) / count);
    return Array.from({ length: count }, (_, index) => ({
        width: columnWidth,
        paddingEnd: index === count - 1 ? 0 : gap,
    }));
}

function lastFootnoteColumn(fragments: IDocumentSkeletonNote[]) {
    const page = fragments[fragments.length - 1]?.page;
    const columns = page?.sections[page.sections.length - 1]?.columns ?? [];
    for (let index = columns.length - 1; index >= 0; index--) {
        const line = columns[index].lines[columns[index].lines.length - 1];
        if (line) {
            const endsParagraph = line.divides.some((divide) => divide.glyphGroup.some((glyph) => glyph.streamType === DataStreamTreeTokenType.PARAGRAPH));
            const afterSpacing = endsParagraph ? Math.max(0, line.spaceBelowApply - line.marginBottom) : 0;
            return { index, top: line.top + line.lineHeight + afterSpacing };
        }
    }
    return { index: 0, top: 0 };
}

/** Uses the body layout pipeline, including table row splitting and inline drawing metrics. */
export function layoutNoteBody(
    parentContext: ILayoutContext,
    reference: INoteReferenceLayout,
    sectionConfig: ISectionBreakConfig,
    constraints: INoteLayoutConstraints
): IDocumentSkeletonPage[] {
    const viewModel = parentContext.viewModel.getNoteTreeMap().get(reference.noteId);
    if (!viewModel) {
        return [];
    }
    return layoutNoteSegment(parentContext, viewModel, reference, sectionConfig, constraints);
}

export function layoutNoteSegment(
    parentContext: ILayoutContext,
    viewModel: DocumentViewModel,
    reference: Pick<INoteReferenceLayout, 'noteId' | 'sectionId'> & { label?: string },
    sectionConfig: ISectionBreakConfig,
    constraints: INoteLayoutConstraints
): IDocumentSkeletonPage[] {
    const skeleton = getNullSkeleton();
    const { skeHeaders, skeFooters, skeListLevel, drawingAnchor } = skeleton;
    const dataModel = viewModel.getDataModel();
    const config: ISectionBreakConfig = {
        ...sectionConfig,
        documentCompatibilityPolicy: sectionConfig.documentCompatibilityPolicy ??
            getDocumentCompatibilityPolicy(parentContext.dataModel.documentStyle.documentFlavor),
        sectionId: reference.sectionId,
        pageSize: { width: constraints.width, height: constraints.continuationPageHeight },
        marginTop: 0,
        marginBottom: 0,
        marginLeft: 0,
        marginRight: 0,
        marginHeader: 0,
        marginFooter: 0,
        headerIds: {},
        footerIds: {},
        headerTreeMap: new Map(),
        footerTreeMap: new Map(),
        columnProperties: createFootnoteColumns(sectionConfig, constraints.width, constraints.columnCount ?? 1),
        columnSeparatorType: ColumnSeparatorType.NONE,
        gridType: GridType.DEFAULT,
        drawings: dataModel.drawings,
        lists: { ...sectionConfig.lists, ...dataModel.lists },
        documentTextStyle: reference.label == null ? sectionConfig.documentTextStyle : { ...sectionConfig.documentTextStyle, fs: 10 },
    };
    const ctx: ILayoutContext = {
        ...parentContext,
        viewModel,
        dataModel,
        skeleton,
        skeletonResourceReference: { skeHeaders, skeFooters, skeListLevel, drawingAnchor },
        noteReferences: undefined,
        footnoteLayout: undefined,
        endnoteLayout: undefined,
        noteLabel: parentContext.dataModel.getSnapshot().notes?.[reference.noteId]?.showBodyReference === false
            ? undefined
            : reference.label,
        noteReferenceTextStyle: parentContext.dataModel.getSnapshot().notes?.[reference.noteId]?.referenceTextStyle,
        noteSegmentId: reference.noteId,
        footnoteFirstColumn: constraints.firstColumn,
        deferSlicedTableLayout: undefined,
        layoutStartPointer: { [reference.noteId]: null },
        isDirty: false,
        floatObjectsCache: new Map(),
        paragraphConfigCache: new Map(),
        sectionBreakConfigCache: new Map(),
        paragraphsOpenNewPage: new Set(),
    };
    const firstPage = createSkeletonPage(ctx, {
        ...config,
        pageSize: { width: constraints.width, height: Math.max(1, constraints.firstPageHeight) },
    }, ctx.skeletonResourceReference);
    firstPage.type = DocumentSkeletonPageType.NOTE;
    firstPage.segmentId = reference.noteId;
    for (let index = 0; index < (constraints.firstColumn?.index ?? 0); index++) {
        firstPage.sections[0].columns[index].isFull = true;
    }
    try {
        const pages = dealWithSection(ctx, viewModel, viewModel.getChildren()[0], firstPage, config, null).pages;
        for (const page of pages) {
            page.type = DocumentSkeletonPageType.NOTE;
            page.segmentId = reference.noteId;
        }
        updateBlockIndex(pages, -1, config.documentCompatibilityPolicy);
        updateInlineDrawingCoordsAndBorder(ctx, pages);
        expandCellPageHeightForInlineDrawings(pages);
        expandCellPageHeightForFlowTables(pages);
        // Word reserves the final paragraph's after-spacing inside a note story.
        for (const page of pages) {
            for (const section of page.sections) {
                for (const column of section.columns) {
                    const line = column.lines[column.lines.length - 1];
                    if (line?.divides.some((divide) => divide.glyphGroup.some((glyph) => glyph.streamType === DataStreamTreeTokenType.PARAGRAPH))) {
                        page.height = Math.max(page.height, section.top + line.top + line.lineHeight +
                            Math.max(0, line.spaceBelowApply - line.marginBottom));
                    }
                }
            }
        }
        return pages;
    } finally {
        // Text run offsets are local to a segment. Never reuse a note's font cache for body text.
        clearFontCreateConfigCache();
    }
}

interface IPageFootnoteLayout {
    fragments: IDocumentSkeletonNote[];
    following: IDocumentSkeletonNote[][];
    referenceIds: string[];
    height: number;
    separator?: IDocumentSkeletonPage;
    notice?: IDocumentSkeletonPage;
    /**
     * Where the body of the page ends when a note that starts on the page continues on the next one: Word puts
     * nothing of the body below the line that holds its reference, even when the note leaves room there.
     */
    bodyEnd?: number;
}

/** Whether a note that starts in the plan's notes continues on the next page. */
function continuesOwnNote(plan: IPageFootnoteLayout): boolean {
    return plan.following[0]?.some((fragment) => fragment.continued && plan.referenceIds.includes(fragment.noteId) &&
        plan.fragments.some((own) => !own.continued && own.noteId === fragment.noteId)) === true;
}

/**
 * Where the body of the page may end with the plan's notes below it. A note that continues takes the whole page below
 * the line of its reference (its last line may go up to 2 px further, as lines do), and the body ends with that line.
 */
function planBodyLimit(contentHeight: number, plan: IPageFootnoteLayout): number {
    return plan.bodyEnd ?? contentHeight - plan.height;
}

type FootnoteDecorationKind = IDocumentSkeletonNoteDecoration['kind'];

export function defaultSeparatorBody(continued: boolean): IDocumentNoteDecoration {
    return {
        dataStream: '\uFFFC\r\n',
        paragraphs: [{ paragraphId: 'separator', startIndex: 1, paragraphStyle: { spaceAbove: { v: 0 }, spaceBelow: { v: 0 } } }],
        sectionBreaks: [{ sectionId: 'separator', startIndex: 2 }],
        customRanges: [{
            rangeId: 'separator',
            rangeType: CustomRangeType.FIELD,
            startIndex: 0,
            endIndex: 0,
            wholeEntity: true,
            properties: { fieldType: continued ? 'FOOTNOTE_CONTINUATION_SEPARATOR' : 'FOOTNOTE_SEPARATOR' },
        }],
    };
}

function footnoteContentHeight(fragments: IDocumentSkeletonNote[]): number {
    return fragments.reduce((height, fragment) => Math.max(height, fragment.page.height), 0);
}

function collectPageReferenceGlyphs(page: IDocumentSkeletonPage): IDocumentSkeletonGlyph[] {
    const result: IDocumentSkeletonGlyph[] = [];
    for (const section of page.sections) {
        for (const column of section.columns) {
            for (const line of column.lines) {
                for (const divide of line.divides) {
                    result.push(...divide.glyphGroup.filter((glyph) => glyph.noteId));
                }
            }
        }
    }
    for (const table of page.skeTables.values()) {
        result.push(...table.rows.flatMap((row) => row.cells.flatMap(collectPageReferenceGlyphs)));
    }
    for (const group of page.skeColumnGroups.values()) {
        result.push(...group.columns.flatMap((column) => collectPageReferenceGlyphs(column.page)));
    }
    return result;
}

function collectPageReferenceIds(page: IDocumentSkeletonPage): string[] {
    return [...new Set(collectPageReferenceGlyphs(page).map((glyph) => glyph.noteId!))];
}

function collectCellReferenceLines(page: IDocumentSkeletonPage, offset = 0): { top: number; bottom: number; glyphs: IDocumentSkeletonGlyph[] }[] {
    const lines = page.sections.flatMap((section) => section.columns.flatMap((column) => column.lines.flatMap((line) => {
        const glyphs = line.divides.flatMap((divide) => divide.glyphGroup.filter((glyph) => glyph.noteId));
        const top = offset + page.marginTop + section.top + line.top;
        return glyphs.length > 0 ? [{ top, bottom: top + line.lineHeight + page.marginBottom, glyphs }] : [];
    })));
    for (const table of page.skeTables.values()) {
        for (const row of table.rows) {
            lines.push(...row.cells.flatMap((cell) => collectCellReferenceLines(cell, offset + page.marginTop + table.top + row.top)));
        }
    }
    return lines;
}

/**
 * How high a part of the row must be to hold the first line of one of its cells, with the cell's margins (0 when a
 * cell starts with something else than a line), as the table layout measures it when a row moves whole.
 */
function rowFirstLineHeight(row: IDocumentSkeletonTable['rows'][number]): number {
    let height = Number.POSITIVE_INFINITY;
    for (const cell of row.cells) {
        const line = cell.sections[0]?.columns[0]?.lines[0];
        if (!line || cell.skeTables.size > 0) {
            return 0;
        }
        height = Math.min(height, (cell.marginTop ?? 0) + line.top + line.lineHeight + (cell.marginBottom ?? 0));
    }
    return Number.isFinite(height) ? height : 0;
}

/** How many lines of a cell page end within the given height, those of the tables inside it included. */
function countCellLinesWithin(page: IDocumentSkeletonPage, height: number, offset = 0): number {
    let count = 0;
    for (const section of page.sections) {
        for (const column of section.columns) {
            for (const line of column.lines) {
                if (offset + page.marginTop + section.top + line.top + line.lineHeight <= height + 0.01) {
                    count++;
                }
            }
        }
    }
    for (const table of page.skeTables.values()) {
        for (const row of table.rows) {
            count += row.cells.reduce((total, cell) => total + countCellLinesWithin(cell, height, offset + page.marginTop + table.top + row.top), 0);
        }
    }
    return count;
}

/** A line of a table row that holds note references, in the coordinates of its cell page. */
export interface INoteReferenceLine {
    top: number;
    /** Where the line ends, with the bottom margin of its cell. */
    bottom: number;
    /** The height a row fragment needs to end just before the line, with the bottom margin of its cell. */
    split: number;
    /** The line opens its cell page: no fragment can end before it. */
    atTop: boolean;
    glyphs: IDocumentSkeletonGlyph[];
}

/**
 * How high the part of a table row on this page can be, given the lines of the row that hold note references (sorted
 * by where they end) and a measurement of the page's notes for the lines it keeps: the body limit those notes leave
 * and the notes the page postpones. The row ends after as many reference lines as can keep their notes on the page,
 * below the text (a note never starts on a page after its reference): the lines are tried from the last one up, as
 * Word fills the page, and a note that has no room, or that a continuing note postpones, sends its reference line
 * and those after it on. A reference line that opens the fragment stays whatever happens to its note.
 */
export function findNoteRowCapacity(
    lines: readonly INoteReferenceLine[],
    fullBottom: number,
    measure: (keptLines: readonly INoteReferenceLine[], bottom: number) => { limit: number; deferredIds: ReadonlySet<string> },
    movable = false
): number {
    // Lines that open their cell page stay whatever happens to their notes (they would continue on the next page),
    // unless the row can move: its part below other content of the page can go to the next page whole.
    let fixed = 0;
    if (!movable) {
        while (fixed < lines.length && lines[fixed].atTop) {
            fixed++;
        }
    }
    for (let kept = lines.length; kept >= fixed; kept--) {
        const keptLines = lines.slice(0, kept);
        const bottom = kept === lines.length ? fullBottom : (keptLines[kept - 1]?.bottom ?? 0);
        const { limit, deferredIds } = measure(keptLines, bottom);
        const deferred = keptLines.some((line) => (movable || !line.atTop) && line.glyphs.some((glyph) => deferredIds.has(glyph.noteId!)));
        // A note that continues ends the body with its reference line: the limit is that line's bottom, measured back
        // from the page, and rounding must not lose it (the verifier's probe v8r-lunga-40: 16.00999999999995 for 16.01).
        if (kept === fixed || (!deferred && limit >= bottom - 0.01)) {
            const capacity = Math.min(limit, lines[kept]?.split ?? Number.POSITIVE_INFINITY);
            return Math.max(1, kept === fixed && kept > 0 ? Math.max(capacity, bottom) : capacity);
        }
    }
    return 1;
}

/** Per-layout state, shared by synchronous and sliced body pagination. */
export class DocumentFootnoteLayout {
    private readonly _pages = new WeakMap<IDocumentSkeletonPage, IPageFootnoteLayout>();
    private readonly _inherited = new WeakMap<IDocumentSkeletonPage, IPageFootnoteLayout>();
    private readonly _references = new Map<string, INoteReferenceLayout>();
    private readonly _referencePageBreaks = new Map<number, Map<number, number>>();
    private readonly _decorations = new Map<string, IDocumentSkeletonPage | undefined>();

    constructor(private readonly _ctx: ILayoutContext) {
        for (const reference of _ctx.noteReferences?.values() ?? []) {
            if (reference.type === 'footnote') {
                this._references.set(reference.noteId, reference);
            }
        }
    }

    updateReferenceGlyphs(
        page: IDocumentSkeletonPage,
        glyphs: IDocumentSkeletonGlyph[],
        config: ISectionBreakConfig,
        paragraph: IParagraphConfig
    ): void {
        if (!glyphs.some((glyph) => glyph.noteId) || page.type !== DocumentSkeletonPageType.BODY || page.segmentId !== '') {
            return;
        }
        this._updateReferenceLabels(page, glyphs);
        for (const glyph of glyphs) {
            const reference = glyph.noteId ? this._references.get(glyph.noteId) : undefined;
            if (!reference || reference.properties.restart !== 'eachPage' || reference.number == null) {
                continue;
            }
            const label = reference.label;
            if (glyph.content !== label) {
                const fontConfig = getFontConfigFromLastGlyph(glyph, config, paragraph.paragraphStyle ?? {});
                Object.assign(glyph, createSkeletonLetterGlyph(label, fontConfig), {
                    raw: '\uFFFC',
                    count: 1,
                    noteId: reference.noteId,
                });
            }
        }
    }

    private _updateReferenceLabels(page: IDocumentSkeletonPage, incoming: IDocumentSkeletonGlyph[]): boolean {
        const ids = [...new Set([...collectPageReferenceIds(page), ...incoming.map((glyph) => glyph.noteId!)])]
            .filter((id) => this._references.has(id))
            .sort((left, right) => this._references.get(left)!.referenceIndex - this._references.get(right)!.referenceIndex);
        let precedingCount = 0;
        let changed = false;
        for (const id of ids) {
            const reference = this._references.get(id)!;
            if (reference.number == null) {
                continue;
            }
            if (reference.properties.restart === 'eachPage') {
                const number = (reference.properties.startNumber ?? 1) + precedingCount;
                const label = formatNoteNumber(number, reference.properties.numberFormat);
                changed ||= reference.label !== label;
                reference.label = label;
            }
            precedingCount++;
        }
        if (changed) {
            this._pages.delete(page);
        }
        return changed;
    }

    getTableBodyLimit(
        page: IDocumentSkeletonPage,
        currentPages: IDocumentSkeletonPage[],
        config: ISectionBreakConfig,
        table: IDocumentSkeletonTable,
        keepReferences = false
    ): number {
        const references = table.rows.flatMap((row) => row.cells.flatMap(collectPageReferenceGlyphs));
        return this.getBodyLimit(page, currentPages, config, table.top + table.height, references, keepReferences);
    }

    /** Whether the page holds the rest of notes that the page before it could not finish. */
    inheritsNotes(page: IDocumentSkeletonPage, currentPages: IDocumentSkeletonPage[], config: ISectionBreakConfig): boolean {
        if (page.type !== DocumentSkeletonPageType.BODY || page.segmentId !== '') {
            return false;
        }
        return this._getInherited(page, currentPages, config).fragments.length > 0;
    }

    /** Measure table slices without publishing speculative notes on body pages. */
    createTablePagination(page: IDocumentSkeletonPage, config: ISectionBreakConfig, availableHeight: number) {
        if (page.type !== DocumentSkeletonPageType.BODY || page.segmentId !== '') {
            return undefined;
        }
        let current = { ...page, skeTables: new Map(page.skeTables) };
        const pages = [current];
        this._inherited.set(current, this._getInherited(page, this._ctx.skeleton.pages, config));
        let top = page.pageHeight - page.marginTop - page.marginBottom - availableHeight;
        return {
            renumberWholeRow: (table: IDocumentSkeletonTable, rows: IDocumentSkeletonTable['rows'], remainingHeight: number): boolean => {
                if (!rows.some((row) => row.cells.some((cell) => collectPageReferenceGlyphs(cell)
                    .some((glyph) => this._references.get(glyph.noteId!)?.properties.restart === 'eachPage')))) {
                    return false;
                }
                let target = { ...current, skeTables: new Map(current.skeTables) };
                const targetPages = [target];
                this._inherited.set(target, this._getInherited(current, pages, config));
                const accepted = this._pages.get(current);
                if (accepted) {
                    this._pages.set(target, accepted);
                }
                let offset = top + table.height;
                let preceding = table.rows.flatMap((row) => row.cells.flatMap(collectPageReferenceGlyphs));
                let changed = false;
                for (const row of rows) {
                    const glyphs = row.cells.flatMap(collectPageReferenceGlyphs);
                    let previousPlan = this._pages.get(target) ?? this._getInherited(target, targetPages, config);
                    this._updateReferenceLabels(target, [...preceding, ...glyphs]);
                    let limit = this.getBodyLimit(target, targetPages, config, offset + row.height, [...preceding, ...glyphs]);
                    const tolerance = (config.documentCompatibilityPolicy ?? getDocumentCompatibilityPolicy()).table.rowOverflowTolerance;
                    const requiresNewPage = row !== rows[0] || remainingHeight <= 0 || row.height > remainingHeight + tolerance;
                    if (requiresNewPage || offset + row.height > limit + 0.01) {
                        do {
                            this._pages.set(target, previousPlan);
                            target = createSkeletonPage(this._ctx, config, this._ctx.skeletonResourceReference, getNextSkeletonPageNumber(target, config));
                            targetPages.push(target);
                            offset = 0;
                            preceding = [];
                            previousPlan = this._getInherited(target, targetPages, config);
                            this._updateReferenceLabels(target, glyphs);
                            limit = this.getBodyLimit(target, targetPages, config, row.height, glyphs);
                        } while (row.height > limit + 0.01 && this._getInherited(target, targetPages, config).fragments.length > 0);
                    }
                    changed ||= glyphs.some((glyph) => glyph.content !== this._references.get(glyph.noteId!)?.label);
                    offset += row.height;
                    preceding.push(...glyphs);
                }
                return changed;
            },
            /**
             * The heights of the parts of a row that its notes allow, page by page. borders are the heights the table's
             * outer borders add to a slice that ends with the first part and with a later one: the notes are measured
             * below the slice as it is laid out on the page.
             */
            measureRow: (table: IDocumentSkeletonTable, rows: IDocumentSkeletonTable['rows'], previousHeights?: readonly number[], startsOnNextPage = false, borders: readonly [number, number] = [0, 0]): number[] | 'next-page' | undefined => {
                let measurePage = { ...current, skeTables: new Map(current.skeTables) };
                const measurePages = [measurePage];
                this._inherited.set(measurePage, this._getInherited(current, pages, config));
                const accepted = this._pages.get(current);
                if (accepted) {
                    this._pages.set(measurePage, accepted);
                }
                if (startsOnNextPage) {
                    // The row cannot start in what is left of the current page (its minimum height does not fit):
                    // its first fragment begins the next page, which it measures against.
                    measurePage = createSkeletonPage(this._ctx, config, this._ctx.skeletonResourceReference, getNextSkeletonPageNumber(measurePage, config));
                    measurePages.push(measurePage);
                }
                const heights = [...(previousHeights ?? [])];
                let changed = false;
                for (let index = 0; index < rows.length; index++) {
                    const row = rows[index];
                    const onCurrentPage = index === 0 && !startsOnNextPage;
                    const offset = onCurrentPage ? top + table.height : 0;
                    const border = index === 0 ? borders[0] : borders[1];
                    const contentHeight = measurePage.pageHeight - measurePage.marginTop - measurePage.marginBottom;
                    const available = heights[index] ?? contentHeight - offset;
                    const lines = row.cells
                        // A fragment that ends before a line also keeps the bottom margin of the line's cell.
                        .flatMap((cell) => collectCellReferenceLines(cell).map((line) => ({
                            ...line,
                            cell,
                            split: line.top + (cell.marginBottom ?? 0),
                            atTop: line.top <= (cell.marginTop ?? 0) + 0.01,
                        })))
                        .filter((line) => line.bottom <= available)
                        .sort((left, right) => left.bottom - right.bottom);
                    const incoming = lines.flatMap((line) => line.glyphs);
                    if (onCurrentPage) {
                        incoming.push(...table.rows.flatMap((item) => item.cells.flatMap(collectPageReferenceGlyphs)));
                    }
                    // A cell is shaped before its physical page is known. Rebuild the
                    // row when its page-local marker changes, including marker width.
                    changed = this._updateReferenceLabels(measurePage, incoming) || changed;
                    const preceding = onCurrentPage ? incoming.slice(lines.reduce((count, line) => count + line.glyphs.length, 0)) : [];
                    const capacity = findNoteRowCapacity(
                        lines,
                        lines[lines.length - 1]?.bottom ?? Math.min(row.height, available),
                        (keptLines, bottom) => {
                            const limit = this.getBodyLimit(measurePage, measurePages, config, offset + bottom + border, [...keptLines.flatMap((line) => line.glyphs), ...preceding]) - offset - border;
                            const deferredIds = new Set(this._pages.get(measurePage)?.following.flatMap((fragments) => fragments
                                .filter((fragment) => !fragment.continued)
                                .map((fragment) => fragment.noteId)));
                            return { limit, deferredIds };
                        },
                        onCurrentPage && offset > 0.01
                    );
                    // Word does not end the first part of a row, below other content of the page, with a single line
                    // of a cell that holds a note reference there: the row goes to the next page (probes
                    // righe-minime-note and righe-orfane; the Mosaico of Golden Group, page 183). A single line without
                    // a reference, or two lines with one, end the part (righe-orfane D1, D3, D5), and so does a single
                    // line whose note starts below it and continues on the next page: Word ends the page with that
                    // line (the verifier's probes v8r-lunga-40 and -44, a note of 20 lines).
                    const partHeight = Math.min(capacity, row.height);
                    // A part below other content of the page that cannot hold the first line of its cells holds
                    // nothing: the row goes to the next page whole, as Word moves it (the verifier's probes v8r-prima-40,
                    // -terza-40, -ultima-40: the row after a note's page went there with one line, and the rest to a
                    // page more).
                    if (onCurrentPage && offset > 0.01 && row.height > capacity + 0.01 && capacity + 0.01 < rowFirstLineHeight(row)) {
                        return 'next-page';
                    }
                    if (onCurrentPage && offset > 0.01 && (rows.length > 1 || row.height > capacity + 0.01)) {
                        const opensEmpty = lines.some((line) => line.atTop && line.bottom > partHeight + 0.01);
                        const kept = lines.filter((line) => line.bottom <= partHeight + 0.01);
                        const single = kept.some((line) => countCellLinesWithin(line.cell, partHeight) < 2);
                        if (opensEmpty || (single && !this._partContinuesNote(measurePage, measurePages, config, offset + partHeight + border, [...kept.flatMap((line) => line.glyphs), ...preceding]))) {
                            // Its first part would hold nothing, or a single line with a reference.
                            return 'next-page';
                        }
                    }
                    if (row.height > capacity + 0.01 && capacity + 0.01 < available) {
                        heights[index] = capacity;
                        changed = true;
                        // The pages after this one were measured with the lines this page now gives them: they are
                        // measured again, from their whole height, once the row is laid out with this one (Word fills
                        // each page with what the notes before it leave).
                        heights.length = index + 1;
                        break;
                    } else {
                        heights[index] = available;
                    }
                    measurePage = createSkeletonPage(this._ctx, config, this._ctx.skeletonResourceReference, getNextSkeletonPageNumber(measurePage, config));
                    measurePages.push(measurePage);
                }
                return changed ? heights : undefined;
            },
            /**
             * What is left of the page for rows and the slice's borders after the row, or undefined when the notes do
             * not let the row end the slice there; border is the height the outer borders add to the slice if it ends
             * with the row, so that the notes are measured below the slice as it is laid out.
             */
            append: (table: IDocumentSkeletonTable, row: IDocumentSkeletonTable['rows'][number], border = 0): number | undefined => {
                const previous = this._pages.get(current);
                const incoming = [...table.rows, row].flatMap((item) => item.cells.flatMap(collectPageReferenceGlyphs));
                const bottom = top + table.height + row.height;
                const limit = this.getBodyLimit(current, pages, config, bottom + border, incoming);
                if (bottom + border <= limit + 0.01) {
                    return Math.max(0, limit - bottom);
                }
                if (previous) {
                    this._pages.set(current, previous);
                } else {
                    this._pages.set(current, this._getInherited(current, pages, config));
                }
                return undefined;
            },
            hasContinuation: () => this._getInherited(current, pages, config).fragments.length > 0,
            nextPage: () => {
                current = createSkeletonPage(this._ctx, config, this._ctx.skeletonResourceReference, getNextSkeletonPageNumber(current, config));
                pages.push(current);
                top = 0;
            },
        };
    }

    /**
     * Where the body of the page may end with the notes of its references below it, or -1 when a note cannot start on
     * the page: the content with its reference has to go to the next page. With keepReferences the content stays
     * whatever happens (nothing comes before it on the page, so the next page would be the same): the notes that
     * cannot start on the page start on the next one.
     */
    getBodyLimit(
        page: IDocumentSkeletonPage,
        currentPages: IDocumentSkeletonPage[],
        config: ISectionBreakConfig,
        bodyBottom: number,
        incoming: IDocumentSkeletonGlyph[] = [],
        keepReferences = false
    ): number {
        const contentHeight = page.pageHeight - page.marginTop - page.marginBottom;
        if (page.type !== DocumentSkeletonPageType.BODY || page.segmentId !== '' || !Number.isFinite(contentHeight)) {
            return contentHeight;
        }
        const inherited = this._getInherited(page, currentPages, config);
        const ids = new Set(collectPageReferenceIds(page));
        for (const glyph of incoming) {
            if (glyph.noteId) {
                ids.add(glyph.noteId);
            }
        }
        const referenceIds = [...ids].filter((id) => this._references.has(id)).sort((left, right) =>
            (this._references.get(left)?.referenceIndex ?? 0) - (this._references.get(right)?.referenceIndex ?? 0));
        const previous = this._pages.get(page);
        if (previous && previous.referenceIds.length === referenceIds.length &&
            previous.referenceIds.every((id, index) => id === referenceIds[index])) {
            this._position(page, previous, bodyBottom);
            return planBodyLimit(contentHeight, previous);
        }

        const firstReference = this._references.get(inherited.fragments[0]?.noteId ?? referenceIds[0]);
        const requestedCount = firstReference?.properties.columnCount ?? 0;
        const columnCount = requestedCount === 0 ? Math.max(1, config.columnProperties?.length ?? 1) : requestedCount;
        const separator = this._layoutDecoration(this._getSeparatorKind(inherited.fragments.some((fragment) => fragment.continued)), page, config);
        const available = Math.max(1, contentHeight - bodyBottom - (separator?.height ?? 0));
        let plan = this._layoutPageNotes(page, config, referenceIds, inherited, available, columnCount, separator, false, keepReferences);
        if (!plan) {
            return -1;
        }
        // Balance all notes on this page together; a continuing note already uses
        // the full available height and must retain its continuation boundary.
        if (columnCount > 1 && plan.following.length === 0 && inherited.fragments.length === 0 && plan.fragments.length > 0) {
            let low = 1;
            let high = Math.max(1, footnoteContentHeight(plan.fragments));
            for (let attempt = 0; attempt < 12 && high - low > 0.25; attempt++) {
                const middle = (low + high) / 2;
                const candidate = this._layoutPageNotes(page, config, referenceIds, inherited, middle, columnCount, separator, false, keepReferences);
                if (candidate && candidate.following.length === 0) {
                    plan = candidate;
                    high = middle;
                } else {
                    low = middle;
                }
            }
        }
        // Word ends the page's body with the line of a reference whose note continues on the next page (probe
        // note-continuate: plain paragraphs and a table row that would fit; the Mosaico of Golden Group, page 263).
        if (continuesOwnNote(plan)) {
            plan.bodyEnd = bodyBottom;
        }
        this._pages.set(page, plan);
        this._position(page, plan, bodyBottom);
        return planBodyLimit(contentHeight, plan);
    }

    /** Whether, with the body ending at bodyBottom and these references, a note that starts on the page continues. */
    private _partContinuesNote(
        page: IDocumentSkeletonPage,
        currentPages: IDocumentSkeletonPage[],
        config: ISectionBreakConfig,
        bodyBottom: number,
        incoming: IDocumentSkeletonGlyph[]
    ): boolean {
        this.getBodyLimit(page, currentPages, config, bodyBottom, incoming);
        const plan = this._pages.get(page);
        return plan != null && continuesOwnNote(plan);
    }

    /**
     * Whether a note that the table references would continue on the next page if the table were laid out whole on
     * the page: Word ends the page with the line of its reference, so the table is laid out row by row. Measured on
     * a copy of the page, whose notes stay as they are.
     */
    tableContinuesNote(
        page: IDocumentSkeletonPage,
        currentPages: IDocumentSkeletonPage[],
        config: ISectionBreakConfig,
        table: IDocumentSkeletonTable
    ): boolean {
        if (page.type !== DocumentSkeletonPageType.BODY || page.segmentId !== '') {
            return false;
        }
        const references = table.rows.flatMap((row) => row.cells.flatMap(collectPageReferenceGlyphs));
        if (!references.some((glyph) => glyph.noteId && this._references.has(glyph.noteId))) {
            return false;
        }
        const probe = { ...page, skeTables: new Map(page.skeTables) };
        this._inherited.set(probe, this._getInherited(page, currentPages, config));
        const accepted = this._pages.get(page);
        if (accepted) {
            this._pages.set(probe, accepted);
        }
        this.getBodyLimit(probe, [probe], config, table.top + table.height, references);
        const plan = this._pages.get(probe);
        return plan != null && continuesOwnNote(plan);
    }

    finish(pages: IDocumentSkeletonPage[], config: ISectionBreakConfig): void {
        let previous = pages[pages.length - 1];
        if (!previous) {
            return;
        }
        while (this._pages.get(previous)?.following.length) {
            const page = createSkeletonPage(this._ctx, config, this._ctx.skeletonResourceReference, getNextSkeletonPageNumber(previous, config));
            pages.push(page);
            const inherited = this._getInherited(page, pages, config);
            this._pages.set(page, inherited);
            this._position(page, inherited, 0);
            previous = page;
        }
    }

    getReferencePageBreaks(paragraphIndex: number): readonly number[] | undefined {
        const boundaries = this._referencePageBreaks.get(paragraphIndex);
        return boundaries ? [...boundaries.values()] : undefined;
    }

    /** Re-shape moved text while retaining the page boundary chosen by keep/widow control. */
    queueReferenceRelayout(pages: IDocumentSkeletonPage[]): void {
        let anchor = Number.POSITIVE_INFINITY;
        for (const page of pages) {
            if (page.type !== DocumentSkeletonPageType.BODY || page.segmentId !== '') {
                continue;
            }
            this._updateReferenceLabels(page, []);
            for (const section of page.sections) {
                for (const column of section.columns) {
                    const offsets = new Map<number, number>();
                    for (const line of column.lines) {
                        let offset = offsets.get(line.paragraphIndex) ?? 0;
                        for (const glyph of line.divides.flatMap((divide) => divide.glyphGroup)) {
                            const reference = this._references.get(glyph.noteId ?? '');
                            if (reference?.properties.restart === 'eachPage' && glyph.content !== reference.label) {
                                const boundaries = this._referencePageBreaks.get(line.paragraphIndex) ?? new Map<number, number>();
                                boundaries.set(page.pageNumber, reference.referenceIndex - offset);
                                this._referencePageBreaks.set(line.paragraphIndex, boundaries);
                                anchor = Math.min(anchor, line.paragraphIndex);
                            }
                            offset += glyph.count;
                        }
                        offsets.set(line.paragraphIndex, offset);
                    }
                }
            }
        }
        if (!Number.isFinite(anchor)) {
            return;
        }
        // The retained physical-page boundary also covers keep-next predecessors.
        // A second whole-paragraph break would separate them again during the retry.
        for (const paragraph of this._ctx.paragraphsOpenNewPage) {
            if (paragraph >= anchor) {
                this._ctx.paragraphsOpenNewPage.delete(paragraph);
            }
        }
        this._ctx.isDirty = true;
        this._ctx.layoutStartPointer[''] = Math.min(anchor, this._ctx.layoutStartPointer[''] ?? Number.POSITIVE_INFINITY);
    }

    /** Rebuild reservations after keep-lines or widow control relocates shaped lines. */
    reconcilePages(pages: IDocumentSkeletonPage[], config: ISectionBreakConfig): boolean {
        let fits = true;
        for (let index = 0; index < pages.length; index++) {
            const page = pages[index];
            // The first page can inherit a continuation from a paragraph outside this move.
            if (index > 0) {
                this._inherited.delete(page);
            }
            this._pages.delete(page);
            this._updateReferenceLabels(page, []);
            const bodyBottom = page.sections.reduce((bottom, section) => Math.max(
                bottom,
                ...section.columns.map((column) => {
                    const line = column.lines[column.lines.length - 1];
                    return line ? section.top + line.top + line.lineHeight : 0;
                })
            ), 0);
            if (bodyBottom > this.getBodyLimit(page, pages, config, bodyBottom) + 2) {
                fits = false;
            }
        }
        return fits;
    }

    private _layoutPageNotes(
        page: IDocumentSkeletonPage,
        config: ISectionBreakConfig,
        referenceIds: string[],
        inherited: IPageFootnoteLayout,
        available: number,
        columnCount: number,
        separator: IDocumentSkeletonPage | undefined,
        reserveNotice = false,
        keepReferences = false
    ): IPageFootnoteLayout | undefined {
        const plan: IPageFootnoteLayout = {
            fragments: [...inherited.fragments],
            following: inherited.following.map((fragments) => [...fragments]),
            referenceIds,
            height: inherited.height,
            separator,
        };
        const width = page.pageWidth - page.marginLeft - page.marginRight;
        const continuationSeparator = this._layoutDecoration(this._getSeparatorKind(true), page, config);
        const notice = this._layoutDecoration('continuationNotice', page, config);
        const continuationHeight = page.pageHeight - page.marginTop - page.marginBottom -
            (continuationSeparator?.height ?? 0) - (notice?.height ?? 0);
        for (const id of referenceIds) {
            const reference = this._references.get(id);
            if (!reference) {
                continue;
            }
            const delayed = plan.following.length > 0;
            const target = delayed ? plan.following[plan.following.length - 1] : plan.fragments;
            const notePages = layoutNoteBody(this._ctx, reference, config, {
                width,
                firstPageHeight: delayed ? continuationHeight : Math.max(1, available - (reserveNotice ? notice?.height ?? 0 : 0)),
                continuationPageHeight: Math.max(1, continuationHeight),
                columnCount,
                firstColumn: lastFootnoteColumn(target),
            });
            const first = notePages[0];
            if (!first) {
                continue;
            }
            if (first.sections.every((section) => section.columns.every((column) => column.lines.length === 0))) {
                // Keep a new reference with the first line of its note. Otherwise
                // a later paragraph retry can inherit a note whose reference moved.
                if (!delayed && !keepReferences && (plan.fragments.length > 0 ||
                    page.sections.some((section) => section.columns.some((column) => column.lines.length > 0)))) {
                    return undefined;
                }
                this._appendFollowing(plan.following, notePages.slice(1), reference, false);
                continue;
            }
            if (delayed) {
                target.push({ noteId: id, referenceIndex: reference.referenceIndex, continued: false, left: page.marginLeft, top: 0, page: first });
                this._appendFollowing(plan.following, notePages.slice(1), reference, true);
                continue;
            }
            if (first.height > available + 2) {
                if (!keepReferences && (page.sections.some((section) => section.columns.some((column) => column.lines.length > 0)) ||
                    plan.fragments.length > 0)) {
                    return undefined;
                }
                this._appendFollowing(plan.following, notePages, reference, false);
                continue;
            }
            plan.fragments.push({
                noteId: id,
                referenceIndex: reference.referenceIndex,
                continued: false,
                left: page.marginLeft,
                top: 0,
                page: first,
            });
            this._appendFollowing(plan.following, notePages.slice(1), reference, true);
        }
        if (!reserveNotice && notice && plan.following.length > 0) {
            return this._layoutPageNotes(page, config, referenceIds, inherited, available, columnCount, separator, true, keepReferences);
        }
        plan.notice = plan.following.length > 0 ? notice : undefined;
        plan.height = plan.fragments.length === 0
            ? 0
            : (separator?.height ?? 0) +
            footnoteContentHeight(plan.fragments) + (plan.notice?.height ?? 0);
        return plan;
    }

    private _getSeparatorKind(continued: boolean): 'separator' | 'continuationSeparator' {
        if (!continued) {
            return 'separator';
        }
        const style = this._ctx.dataModel.documentStyle;
        // Word 2013 layout reuses the ordinary footnote separator on continuation pages.
        if (getDocumentCompatibilityPolicy(style.documentFlavor).mode === 'traditional' && Number(style.compatibilitySettings?.compatibilityMode) >= 15) {
            return 'separator';
        }
        return 'continuationSeparator';
    }

    private _layoutDecoration(kind: FootnoteDecorationKind, page: IDocumentSkeletonPage, config: ISectionBreakConfig): IDocumentSkeletonPage | undefined {
        const width = page.pageWidth - page.marginLeft - page.marginRight;
        const key = `${kind}:${config.sectionId}:${width}`;
        if (this._decorations.has(key)) {
            return this._decorations.get(key);
        }
        const snapshot = this._ctx.dataModel.getSnapshot();
        const body = snapshot.noteSettings?.footnote?.[kind] ?? (kind === 'continuationNotice' ? undefined : defaultSeparatorBody(kind === 'continuationSeparator'));
        if (!body) {
            this._decorations.set(key, undefined);
            return undefined;
        }
        const model = new DocumentDataModel({ id: `footnote-${kind}`, body, tableSource: body.tableSource, documentStyle: snapshot.documentStyle, styles: snapshot.styles });
        const viewModel = new DocumentViewModel(model);
        try {
            const pages = layoutNoteSegment(this._ctx, viewModel, { noteId: model.getUnitId(), sectionId: config.sectionId }, config, {
                width,
                firstPageHeight: Number.POSITIVE_INFINITY,
                continuationPageHeight: Number.POSITIVE_INFINITY,
            });
            const decoration = pages[0];
            if (decoration) {
                for (const section of decoration.sections) {
                    for (const column of section.columns) {
                        const lastLine = column.lines[column.lines.length - 1];
                        if (lastLine) {
                            decoration.height = Math.max(decoration.height, section.top + lastLine.top + lastLine.lineHeight + Math.max(0, lastLine.spaceBelowApply ?? 0));
                        }
                    }
                }
            }
            this._decorations.set(key, decoration);
            return decoration;
        } finally {
            viewModel.dispose();
            model.dispose();
        }
    }

    private _appendFollowing(
        following: IDocumentSkeletonNote[][],
        pages: IDocumentSkeletonPage[],
        reference: INoteReferenceLayout,
        continued: boolean
    ): void {
        for (let index = 0; index < pages.length; index++) {
            following.push([{
                noteId: reference.noteId,
                referenceIndex: reference.referenceIndex,
                continued: continued || index > 0,
                left: 0,
                top: 0,
                page: pages[index],
            }]);
        }
    }

    private _getInherited(page: IDocumentSkeletonPage, currentPages: IDocumentSkeletonPage[], config: ISectionBreakConfig): IPageFootnoteLayout {
        const existing = this._inherited.get(page);
        if (existing) {
            return existing;
        }
        const pageIndex = currentPages.indexOf(page);
        let previous = pageIndex > 0 ? currentPages[pageIndex - 1] : undefined;
        if (!previous) {
            // The first page of a block's pages: the page before it is among the pages laid out before the block,
            // those of the sections before (the skeleton's) and those of this section so far, which reach the
            // skeleton only when the section ends (a page that a page break opened is the first of the next block).
            const laidOut = this._ctx.skeleton.pages;
            let pages = laidOut;
            if (this._ctx.sectionPages?.length) {
                const known = new Set(laidOut);
                pages = [...laidOut, ...this._ctx.sectionPages.filter((item) => !known.has(item))];
            }
            const index = pages.indexOf(page);
            previous = index > 0 ? pages[index - 1] : index < 0 ? pages[pages.length - 1] : undefined;
            if (previous === page) {
                previous = undefined;
            }
        }
        const previousPlan = previous == null ? undefined : this._pages.get(previous);
        const fragments = previousPlan?.following[0] ?? [];
        const separator = fragments.length > 0
            ? this._layoutDecoration(this._getSeparatorKind(fragments.some((fragment) => fragment.continued)), page, config)
            : undefined;
        const notice = (previousPlan?.following.length ?? 0) > 1 ? this._layoutDecoration('continuationNotice', page, config) : undefined;
        const plan: IPageFootnoteLayout = {
            fragments: fragments.map((fragment) => ({ ...fragment, left: page.marginLeft })),
            following: previousPlan?.following.slice(1) ?? [],
            referenceIds: [],
            height: fragments.length === 0
                ? 0
                : (separator?.height ?? 0) + footnoteContentHeight(fragments) + (notice?.height ?? 0),
            separator,
            notice,
        };
        this._inherited.set(page, plan);
        return plan;
    }

    private _position(page: IDocumentSkeletonPage, plan: IPageFootnoteLayout, bodyBottom: number): void {
        if (plan.fragments.length === 0) {
            page.notes = page.notes?.filter((note) => this._ctx.dataModel.getSnapshot().notes?.[note.noteId]?.type === 'endnote');
            delete page.footnoteHeight;
            page.noteDecorations = page.noteDecorations?.filter((decoration) => decoration.noteType === 'endnote');
            return;
        }
        const firstReference = this._references.get(plan.fragments[0].noteId);
        const separatorHeight = plan.separator?.height ?? 0;
        const top = firstReference?.properties.position === 'beneathText'
            ? page.marginTop + bodyBottom + separatorHeight
            : page.pageHeight - page.marginBottom - plan.height + separatorHeight;
        const endnotes = page.notes?.filter((note) => this._ctx.dataModel.getSnapshot().notes?.[note.noteId]?.type === 'endnote') ?? [];
        page.notes = [...endnotes, ...plan.fragments.map((fragment) => {
            const positioned = { ...fragment, top, left: page.marginLeft };
            return positioned;
        })];
        page.footnoteHeight = plan.height;
        const decorations: IDocumentSkeletonNoteDecoration[] = [];
        if (plan.separator) {
            decorations.push({ kind: this._getSeparatorKind(plan.fragments.some((fragment) => fragment.continued)), left: page.marginLeft, top: top - separatorHeight, page: plan.separator });
        }
        if (plan.notice) {
            decorations.push({ kind: 'continuationNotice', left: page.marginLeft, top: top + footnoteContentHeight(plan.fragments), page: plan.notice });
        }
        page.noteDecorations = [...page.noteDecorations?.filter((decoration) => decoration.noteType === 'endnote') ?? [], ...decorations];
    }
}
