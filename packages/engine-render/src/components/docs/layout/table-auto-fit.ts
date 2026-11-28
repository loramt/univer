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
    DataStreamTreeNodeType,
    NumberUnitType,
    resolveDocumentParagraphBullet,
    resolveDocumentParagraphStyle,
    TableLayoutType,
    TableSizeType,
} from '@univerjs/core';
import type { ISectionBreakConfig } from '../../../basics/interfaces';
import type { DataStreamTreeNode } from '../view-model/data-stream-tree-node';
import type { DocumentViewModel } from '../view-model/document-view-model';
import type { ILayoutContext } from './tools';
import { getDocumentCompatibilityPolicy, isTraditionalDocumentCompatibility } from '../document-compatibility';
import { shaping } from './block/paragraph/shaping';
import { createNullCellPage } from './model/page';
import { clearFontCreateConfigCache, getNumberUnitValue, resolveParagraphCharacterIndents } from './tools';

/**
 * The preferred widths of an autofit table of Word and of its columns, as the converter of Word documents reads them
 * (tblW, and the tcW of the cells of each column): a share of the width (pct, 0 to 1), a length (dxa, px) or none
 * (auto). With them the columns are divided as Word divides them, by the content of their cells.
 */
export interface IWordAutoFit {
    width: { type: 'pct' | 'dxa' | 'auto'; value: number };
    columns: Array<{ type: 'pct' | 'dxa' | 'auto'; value: number }>;
}

/**
 * The narrowest and the widest each column of a table can be (px, cell padding and indents included): the widest
 * piece of text that cannot break, and all of its text on one line. Null when the table has cells this cannot
 * measure (merged, vertical, with tables or lists inside).
 */
function measureTableColumns(
    ctx: ILayoutContext,
    viewModel: DocumentViewModel,
    tableNode: DataStreamTreeNode,
    sectionBreakConfig: ISectionBreakConfig,
    table: ITable
): { minimums: number[]; maximums: number[]; preferred: Array<number | undefined>; total: number } | null {
    const snapshot = viewModel.getSnapshot?.();
    if (!snapshot || table.tableColumns.length === 0) {
        return null;
    }
    // shortcut: nested, vertical and merged cells keep their authored grid until their intrinsic measures are available.
    if (table.tableRows.some((row) => row.gridBefore || row.gridAfter || row.tableCells.some((cell) =>
        (cell.columnSpan ?? 1) !== 1 || (cell.rowSpan ?? 1) !== 1 || (cell.textDirection && cell.textDirection !== 'lrTb')))) {
        return null;
    }
    const grid = table.tableColumns.map((column) => column.size.width.v);
    const total = grid.reduce((sum, width) => sum + width, 0);
    if (!Number.isFinite(total) || total <= 0 || grid.some((width) => !Number.isFinite(width) || width <= 0)) {
        return null;
    }
    const minimums = grid.map(() => 0);
    const maximums = grid.map(() => 0);
    const preferred: Array<number | undefined> = grid.map(() => undefined);
    try {
        for (const [rowIndex, row] of table.tableRows.entries()) {
            const rowNode = tableNode.children[rowIndex];
            if (!rowNode || row.tableCells.length !== grid.length) {
                return null;
            }
            for (const [column, cell] of row.tableCells.entries()) {
                const section = rowNode.children[column]?.children[0];
                if (!section || section.children.some((paragraph) => paragraph.nodeType !== DataStreamTreeNodeType.PARAGRAPH || paragraph.children.length > 0)) {
                    return null;
                }
                // Grid widths may be too narrow even for padding; measure the authored insets before that clamp.
                const wide = createNullCellPage(ctx, sectionBreakConfig, {
                    ...table,
                    tableColumns: table.tableColumns.map((source) => ({ ...source, size: { ...source.size, width: { ...source.size.width, v: total } } })),
                }, rowIndex, column);
                const padding = wide.page.marginLeft + wide.page.marginRight;
                for (const paragraphNode of section.children) {
                    const paragraph = viewModel.getParagraph(paragraphNode.endIndex);
                    if (paragraph && resolveDocumentParagraphBullet(snapshot, paragraph)) {
                        return null;
                    }
                    clearFontCreateConfigCache();
                    const pieces = shaping(ctx, paragraphNode.content ?? '', viewModel, paragraphNode, wide.sectionBreakConfig);
                    let paragraphPadding = padding;
                    let firstLineIndent = 0;
                    if (paragraph && isTraditionalDocumentCompatibility(sectionBreakConfig.documentCompatibilityPolicy ?? getDocumentCompatibilityPolicy())) {
                        const style = resolveParagraphCharacterIndents(
                            ctx,
                            pieces.flatMap((piece) => piece.glyphs),
                            resolveDocumentParagraphStyle(snapshot.documentStyle, paragraph.paragraphStyle, {
                                styles: snapshot.styles,
                                paragraphStyleId: paragraph.styleId,
                            })
                        );
                        paragraphPadding += getNumberUnitValue(style.indentStart, 0) + getNumberUnitValue(style.indentEnd, 0);
                        const first = getNumberUnitValue(style.indentFirstLine, 0);
                        firstLineIndent = first > 0 ? first : -Math.max(0, getNumberUnitValue(style.hanging, 0));
                    }
                    let maximum = 0;
                    let trailing = 0;
                    for (const [pieceIndex, piece] of pieces.entries()) {
                        let width = 0;
                        let pieceTrailing = 0;
                        for (const glyph of piece.glyphs) {
                            width += glyph.width;
                            if (glyph.raw.trim().length === 0) {
                                pieceTrailing += glyph.width;
                            } else {
                                pieceTrailing = 0;
                            }
                        }
                        minimums[column] = Math.max(minimums[column], width - pieceTrailing + paragraphPadding + (pieceIndex === 0 ? firstLineIndent : 0));
                        maximum += width;
                        trailing = width === pieceTrailing ? trailing + pieceTrailing : pieceTrailing;
                    }
                    maximums[column] = Math.max(maximums[column], maximum - trailing + paragraphPadding + firstLineIndent);
                }
                const width = cell.size?.width;
                let value: number | undefined;
                if (width?.u === NumberUnitType.PERCENT) {
                    value = total * width.v / 100;
                } else if (cell.tcWidthType === 'pct' && width) {
                    value = total * width.v * 15 / 5000;
                } else if ((cell.tcWidthType == null || cell.tcWidthType === 'dxa') && width) {
                    value = width.v;
                }
                if (value != null && Number.isFinite(value) && value > 0) {
                    preferred[column] = Math.max(preferred[column] ?? 0, value);
                }
            }
        }
    } finally {
        clearFontCreateConfigCache();
    }
    return { minimums, maximums, preferred, total };
}

export function resolveTableAutoFitColumns(
    ctx: ILayoutContext,
    viewModel: DocumentViewModel,
    tableNode: DataStreamTreeNode,
    sectionBreakConfig: ISectionBreakConfig,
    table: ITable,
    availableWidth = Number.POSITIVE_INFINITY
): ITable {
    if (table.layout === TableLayoutType.FIXED || !viewModel.getSnapshot?.() || table.tableColumns.length === 0 ||
        (table.size?.type === TableSizeType.SPECIFIED && !table.tableRows.some((row) =>
            row.tableCells.some((cell) => cell.tcWidthType != null)))) {
        return table;
    }
    const measured = measureTableColumns(ctx, viewModel, tableNode, sectionBreakConfig, table);
    if (!measured) {
        return table;
    }
    const { minimums, maximums, preferred, total } = measured;
    const grid = table.tableColumns.map((column) => column.size.width.v);
    const minimum = minimums.map((width) => Math.ceil(width * 15 - 1e-8) / 15);
    const minTotal = minimum.reduce((sum, width) => sum + width, 0);
    const widths = grid.map((_, column) => Math.max(minimum[column], preferred[column] ?? maximums[column]));
    const desired = widths.reduce((sum, width) => sum + width, 0);
    if (desired <= 0) {
        return table;
    }
    // An automatic table's grid is a starting point, not a preferred table width.
    const target = table.size.type === TableSizeType.SPECIFIED
        ? Math.max(total, minTotal)
        : Math.ceil(Math.max(minTotal, Math.min(availableWidth, desired)) * 15 - 1e-8) / 15;
    if (desired > target) {
        const flexible = desired - minTotal;
        for (let column = 0; column < widths.length; column++) {
            widths[column] -= (desired - target) * (widths[column] - minimum[column]) / flexible;
        }
    } else if (desired < target) {
        let weights = widths.map((width, column) => preferred[column] == null ? Math.max(0, maximums[column] - minimum[column]) : 0);
        let weightTotal = weights.reduce((sum, weight) => sum + weight, 0);
        if (weightTotal === 0) {
            weights = widths.map((width, column) => preferred[column] == null ? Math.max(width, 1) : 0);
            weightTotal = weights.reduce((sum, weight) => sum + weight, 0);
        }
        if (weightTotal === 0) {
            weights = widths;
            weightTotal = desired;
        }
        for (let column = 0; column < widths.length; column++) {
            widths[column] += (target - desired) * weights[column] / weightTotal;
        }
    }
    let used = 0;
    const resolved = widths.map((width, column) => {
        const value = column === widths.length - 1 ? target - used : Math.round(width * 15) / 15;
        used += value;
        return value;
    });
    if (resolved.every((width, column) => Math.abs(width - grid[column]) < 0.01)) {
        return table;
    }
    return { ...table, tableColumns: table.tableColumns.map((column, index) => ({
        ...column,
        size: { ...column.size, width: { ...column.size.width, v: resolved[index] } },
    })) };
}

/**
 * The columns of an autofit table of Word (`wordAutoFit`, from the converter of Word documents), divided as Word 365
 * divides them (sonde autofit, autofit2 and vi-larghe-pct of the Mosaico bench):
 * - the table is as wide as it asks (a share of the width available, a length), or, without a width of its own, as
 *   wide as its columns want, at most the width available;
 * - a column in percent takes its share of the table, while the shares add up to 100% (the column after gets none),
 *   a column with a length that length, and the others the width of their text on one line; never less than the
 *   narrowest their text allows (the widest word, with the padding);
 * - when that is wider than the table, the columns without a width give up their room first, in proportion to what
 *   they have above their narrowest, then the columns with a length, then those in percent, in proportion to their
 *   share;
 * - when it is narrower, the columns without a width take the room left in proportion to their widest, or, without
 *   them, the columns with a length in proportion to their length, or those in percent to their share.
 * The table that comes out has those columns and that width; a table this cannot measure keeps the columns the
 * converter computed.
 */
export function resolveWordAutoFitColumns(
    ctx: ILayoutContext,
    viewModel: DocumentViewModel,
    tableNode: DataStreamTreeNode,
    sectionBreakConfig: ISectionBreakConfig,
    table: ITable,
    availableWidth: number
): ITable {
    const autoFit = (table as ITable & { wordAutoFit?: IWordAutoFit }).wordAutoFit;
    if (!autoFit || autoFit.columns.length !== table.tableColumns.length || !(availableWidth > 0)) {
        return table;
    }
    const measured = measureTableColumns(ctx, viewModel, tableNode, sectionBreakConfig, table);
    if (!measured) {
        return table;
    }
    const minimum = measured.minimums.map((width) => Math.ceil(width * 15 - 1e-8) / 15);
    const maximum = measured.maximums.map((width, column) => Math.max(width, minimum[column]));
    const count = minimum.length;
    const { width: tableWidth, columns } = autoFit;
    // the shares of the columns in percent, while they add up to 100%: the column after gets none
    let left = 1;
    const shares = columns.map((column) => {
        if (column.type !== 'pct') {
            return 0;
        }
        const share = Math.max(0, Math.min(column.value, left));
        left -= share;
        return share;
    });
    const desiredFor = (target: number) => columns.map((column, index) => {
        if (column.type === 'pct') {
            return Math.max(minimum[index], shares[index] * target);
        }
        return column.type === 'dxa' && column.value > 0 ? Math.max(minimum[index], column.value) : maximum[index];
    });
    let target = tableWidth.type === 'pct' && tableWidth.value > 0
        ? tableWidth.value * availableWidth
        : tableWidth.type === 'dxa' && tableWidth.value > 0 ? tableWidth.value : 0;
    if (target === 0) {
        // without a width of its own: as wide as its columns want, at most the width available
        const wanted = desiredFor(availableWidth).reduce((sum, width) => sum + width, 0);
        target = Math.min(availableWidth, wanted);
    }
    const minTotal = minimum.reduce((sum, width) => sum + width, 0);
    target = Math.max(target, minTotal);
    const widths = desiredFor(target);
    const kinds = columns.map((column) => column.type === 'pct' ? 'pct' : column.type === 'dxa' && column.value > 0 ? 'dxa' : 'auto');
    let excess = widths.reduce((sum, width) => sum + width, 0) - target;
    if (excess > 0.001) {
        for (const kind of ['auto', 'dxa', 'pct'] as const) {
            const group = widths.map((_, index) => index).filter((index) => kinds[index] === kind);
            const weights = group.map((index) => kind === 'pct' ? shares[index] : widths[index] - minimum[index]);
            const flexible = group.reduce((sum, index) => sum + (widths[index] - minimum[index]), 0);
            const weightTotal = weights.reduce((sum, weight) => sum + weight, 0);
            if (flexible <= 0 || weightTotal <= 0) {
                continue;
            }
            const take = Math.min(excess, flexible);
            group.forEach((index, i) => {
                widths[index] = Math.max(minimum[index], widths[index] - take * weights[i] / weightTotal);
            });
            excess -= take;
            if (excess <= 0.001) {
                break;
            }
        }
    } else if (excess < -0.001) {
        const room = -excess;
        for (const kind of ['auto', 'dxa', 'pct'] as const) {
            const group = widths.map((_, index) => index).filter((index) => kinds[index] === kind);
            const weights = group.map((index) => kind === 'auto' ? maximum[index] : kind === 'dxa' ? widths[index] : shares[index]);
            const weightTotal = weights.reduce((sum, weight) => sum + weight, 0);
            if (!group.length || weightTotal <= 0) {
                continue;
            }
            group.forEach((index, i) => {
                widths[index] += room * weights[i] / weightTotal;
            });
            break;
        }
    }
    const sum = widths.reduce((total, width) => total + width, 0);
    if (count === 0 || !Number.isFinite(sum) || sum <= 0) {
        return table;
    }
    return {
        ...table,
        size: { ...table.size, type: TableSizeType.SPECIFIED, width: { ...table.size?.width, v: sum, u: undefined } },
        tableColumns: table.tableColumns.map((column, index) => ({
            ...column,
            size: { ...column.size, type: TableSizeType.SPECIFIED, width: { ...column.size.width, v: widths[index] } },
        })),
    };
}
