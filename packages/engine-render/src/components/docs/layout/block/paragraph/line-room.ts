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

/** Whether the page compares its lines with their room exactly, as Word does: the body and the cells of Word documents. */
export function comparesExactly(page: IDocumentSkeletonPage, paragraphConfig: IParagraphConfig): boolean {
    return (page.type === DocumentSkeletonPageType.CELL || page.type === DocumentSkeletonPageType.BODY) &&
        isTraditionalDocumentCompatibility(paragraphConfig.documentCompatibilityPolicy);
}

/**
 * The room of a line at the foot of its page or of its part of a table row. On the pages of Word documents a line passes
 * its room as soon as it ends below it, as Word compares them: exactly, with nothing but the error of floating point.
 * Word lays lines out at fractions of a twip (an automatic line of Queens Park at 10 pt is 234.4 twips) and counts the
 * widths of borders in whole twips, which the conversion gives it (probe soglia-twip of the bench, Word 365: the
 * last of 60 exact lines, with a bottom border of 2 to 7 eighths of a point, ends from 3.5 twips before to 3.5 twips
 * past the body; the lines that Word keeps are those whose border, in whole twips, ends within the body; the Mosaico of
 * Golden Group: the lines of cells 0.2, 0.24 and 0.67 twip past their room go to the next page in Word). Elsewhere
 * the layout keeps its tolerance.
 *
 * Every check of the room of a line asks it here, and gets the same answer for the same line: where one check let a
 * line stay and another did not, the line was laid out again without end (fork 8b580ea3a, which kept lines less than
 * a twip past their room: the last line of a paragraph in a cell of the Mosaico, 0.24 twip past, until the stack ran
 * out). Each check has its own name, which the tests use to make two checks disagree on purpose.
 */
export const lineRoom = {
    /** Whether the line passes its room: where the line goes. */
    overflows(page: IDocumentSkeletonPage, paragraphConfig: IParagraphConfig, bottom: number, limit: number): boolean {
        return bottom - limit > (comparesExactly(page, paragraphConfig) ? 1e-6 : LINE_LAYOUT_OVERFLOW_TOLERANCE);
    },

    /** Whether the last line of a paragraph in a part of a table row, with its spacing after, passes its room. */
    cellParagraphEndOverflows(page: IDocumentSkeletonPage, paragraphConfig: IParagraphConfig, bottom: number, limit: number): boolean {
        return lineRoom.overflows(page, paragraphConfig, bottom, limit);
    },
};
