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

/** The mark of the warnings of the layout. The tests of the engine and the benches of the PDF take them as errors. */
export const LAYOUT_WARNING = '[univer layout]';

const MAX_PENDING_WARNINGS = 100;
const pending: string[] = [];

/**
 * Records and prints a warning of the layout: it reached one of its limits on laying the same thing out again and
 * took a decision of its own, so that two checks that disagree end the layout instead of running without end or until
 * the stack runs out. Reaching a limit is a defect to look for: the warning keeps it visible.
 */
export function warnLayout(message: string): void {
    const text = `${LAYOUT_WARNING} ${message}`;
    if (pending.length < MAX_PENDING_WARNINGS) {
        pending.push(text);
    }
    console.warn(text);
}

/** The warnings of the layout since the last call, which it clears. */
export function takeLayoutWarnings(): string[] {
    return pending.splice(0);
}
