import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");
const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

test("generation panel width follows visible controls instead of model-specific constants", () => {
    assert.match(
        indexHtml,
        /var updateAutoWidth = function \(panel\) \{[\s\S]*?Math\.min\(getMinimumWidth\(panel\), viewportWidth\)[\s\S]*?--canvas-panel-auto-width/s,
    );
    assert.doesNotMatch(indexHtml, /width: min\((?:900|860|760)px, calc\(100vw - 32px\)\)/);
});

test("gpt-image-2 toolbar stays on one row and official mode gets enough width", () => {
    assert.match(
        indexHtml,
        /\.canvas-generation-panel:has\(\.canvas-model-select\[title\^="gpt-image-2"\]\) \.canvas-generation-toolbar \{[^}]*height: 48px !important;[^}]*min-height: 48px;[^}]*flex-wrap: nowrap;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel:has\(\.canvas-model-select\[title\^="gpt-image-2"\]\) \.canvas-generation-toolbar > div:first-child \{[^}]*justify-content: space-between;/s,
    );
});

test("generation cost is larger and keeps tabular digits", () => {
    assert.match(
        indexHtml,
        /\.canvas-generation-toolbar \.canvas-generate-button \.bg-pink-100 \{[^}]*font-size: 12px !important;[^}]*font-weight: 800 !important;[^}]*font-variant-numeric: tabular-nums;/s,
    );
});

test("running generation button keeps only one compact status icon", () => {
    assert.match(
        bundle,
        /children:e\?y\.jsxs\(y\.Fragment,\{children:\[y\.jsx\(Ky,\{className:"size-4 animate-spin"\}\),y\.jsx\("span",\{className:"text-xs font-medium",children:"\\u505C\\u6B62"\}\)\]\}\)/,
    );
    assert.doesNotMatch(
        bundle,
        /children:e\?y\.jsxs\(y\.Fragment,\{children:\[y\.jsx\(Ky,[\s\S]{0,180}?y\.jsx\(E7,/,
    );
});
