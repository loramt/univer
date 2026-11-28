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

import { BooleanNumber, DocumentFlavor, SpacingRule } from '@univerjs/core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createParagraphLayoutTestBed } from '../block/paragraph/__tests__/create-paragraph-layout-test-bed';
import { DocumentSkeleton } from '../doc-skeleton';
import { FontCache } from '../shaping-engine/font-cache';

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

describe('lines at the foot of a page of a Word document, in whole twips', () => {
    afterEach(() => vi.restoreAllMocks());

    // Word lays out in whole twips (15 to the pixel): a line that ends less than a twip past the body stays on its
    // page, at a twip and more it goes to the next one (probe soglia-twip of the bench: half a twip stays, one and one
    // and a half go). The body is 900 twips; the three lines end 0.5 and 1.5 twips past it, or exactly at its end. The
    // Mosaico of Golden Group, page 300: an empty paragraph 0.1 twip past the room of its row stays, as in Word.
    it.each([[20, 1], [20 + 0.5 / 45, 1], [20 + 1.5 / 45, 2]])('with lines %s high the three lines take %i page(s)', (lineHeight, pages) => {
        expect(pagesFor(lineHeight)).toBe(pages);
    });
});
