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

import type { IColorStyle, Nullable } from '@univerjs/core';
import { getColorStyle, ThemeService } from '@univerjs/core';

const THEME_COLOR_TOKEN_PATTERN = /^[a-z][a-z0-9-]*\.\d+$/i;

// Theme tokens (`blue.600`, the colour Univer gives the text of a link) are no colour a canvas reads: it would keep
// the previous fill, black in practice. They are resolved against the default theme, as the text of a document has
// the same colour in every theme of the interface.
let defaultThemeColors: Nullable<ThemeService> = null;
function themeColor(token: string): Nullable<string> {
    defaultThemeColors ??= new ThemeService();
    const value = defaultThemeColors.getColorFromTheme<unknown>(token);
    return typeof value === 'string' ? value.toLowerCase() : null;
}

export function getColorStyleForCanvas(color: Nullable<IColorStyle>): Nullable<string> {
    const rgb = color?.rgb;
    if (rgb && /^#?auto$/i.test(rgb.trim())) {
        return undefined;
    }
    if (rgb && THEME_COLOR_TOKEN_PATTERN.test(rgb)) {
        return themeColor(rgb) ?? rgb;
    }

    // Older imported snapshots retain OOXML's bare RRGGBB form. ColorKit
    // treats it as invalid/transparent; normalize only this unambiguous form.
    if (rgb && /^[\da-f]{6}$/i.test(rgb)) {
        return `#${rgb.toLowerCase()}`;
    }

    return getColorStyle(color);
}
