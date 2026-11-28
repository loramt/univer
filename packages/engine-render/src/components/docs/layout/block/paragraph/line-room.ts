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

import type { IDocumentSkeletonPage } from '../../../../../basics/i-document-skeleton-cached';
import type { IParagraphConfig } from '../../../../../basics/interfaces';
import { DocumentSkeletonPageType } from '../../../../../basics/i-document-skeleton-cached';
import { isTraditionalDocumentCompatibility } from '../../../document-compatibility';

/** How far past its room a line of a document that is not Word's may end, in pixels. */
export const LINE_LAYOUT_OVERFLOW_TOLERANCE = 2;

/** Whether the page lays out its lines in whole twips, as Word does: the body and the cells of Word documents. */
export function measuresInTwips(page: IDocumentSkeletonPage, paragraphConfig: IParagraphConfig): boolean {
    return (page.type === DocumentSkeletonPageType.CELL || page.type === DocumentSkeletonPageType.BODY) &&
        isTraditionalDocumentCompatibility(paragraphConfig.documentCompatibilityPolicy);
}

/**
 * Whether a line that ends at bottom passes the room that ends at limit, both from the top of the section, as Word
 * measures them: the room in whole twips, the line without the part of a twip it ends with. Word keeps on its page a
 * line that ends less than a twip past its room (probe soglia-twip of the bench, Word 365: lines exactly at the end
 * of the body and half a twip past it stay, at one twip and more they go to the next page, with five widths of
 * border; the Mosaico of Golden Group, page 300: an empty paragraph 0.1 twip past the room its notes leave stays).
 * The room is rounded to the twip because the conversion gives pixels to the hundredth, a fifth of a twip.
 */
export function overflowsInTwips(bottom: number, limit: number): boolean {
    // 1 px is 15 twips
    return Math.floor(bottom * 15 + 1e-6) > Math.round(limit * 15);
}

/**
 * The room of a line at the foot of its page or of its part of a table row: in whole twips on the pages of Word
 * documents, with the layout tolerance elsewhere. Every check of the room of a line asks it here, and gets the same
 * answer for the same line: where one check let a line stay and another did not, the line was laid out again without
 * end (the Mosaico of Golden Group with the rule of twips alone, fork 8b580ea3a: the last line of a paragraph in a
 * cell, 0.3 twip past its room, until the stack ran out). Each check has its own name, which the tests use to make two
 * checks disagree on purpose.
 */
export const lineRoom = {
    /** Whether the line passes its room: where the line goes. */
    overflows(page: IDocumentSkeletonPage, paragraphConfig: IParagraphConfig, bottom: number, limit: number): boolean {
        return measuresInTwips(page, paragraphConfig)
            ? overflowsInTwips(bottom, limit)
            : bottom - limit > LINE_LAYOUT_OVERFLOW_TOLERANCE;
    },

    /** Whether the last line of a paragraph in a part of a table row, with its spacing after, passes its room. */
    cellParagraphEndOverflows(page: IDocumentSkeletonPage, paragraphConfig: IParagraphConfig, bottom: number, limit: number): boolean {
        return lineRoom.overflows(page, paragraphConfig, bottom, limit);
    },
};
