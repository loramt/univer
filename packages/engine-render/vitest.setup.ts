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

import getCanvasWindow from 'jest-canvas-mock/lib/window';
import { afterEach, vi } from 'vitest';
import { takeLayoutWarnings } from './src/components/docs/layout/layout-warnings';

const apis = [
    'Path2D',
    'CanvasGradient',
    'CanvasPattern',
    'CanvasRenderingContext2D',
    'DOMMatrix',
    'ImageData',
    'TextMetrics',
    'ImageBitmap',
    'createImageBitmap',
] as const;

const canvasWindow = getCanvasWindow({ document: window.document });

apis.forEach((api) => {
    vi.stubGlobal(api, canvasWindow[api]);
});

// A warning of the layout is a bound it reached on laying the same thing out again, where it took a decision of its
// own: a defect, which fails the test that met it. The tests of the bounds take their warnings themselves.
afterEach(() => {
    const warnings = takeLayoutWarnings();
    if (warnings.length > 0) {
        throw new Error(`The layout warned: ${warnings.join('; ')}`);
    }
});
