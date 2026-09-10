import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");

test("generation panels use independent automatic and manually locked dimensions", () => {
    assert.match(
        indexHtml,
        /\.canvas-generation-panel \{[^}]*width: min\(var\(--canvas-panel-auto-width, 660px\), calc\(100vw - 32px\)\) !important;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel\[data-canvas-resize-width-ready="true"\] \{[^}]*width: var\(--canvas-panel-resize-width\) !important;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel\[data-canvas-resize-height-ready="true"\] \{[^}]*height: var\(--canvas-panel-resize-height\) !important;/s,
    );
    assert.doesNotMatch(indexHtml, /width: min\((?:740|760|860|900)px, calc\(100vw - 32px\)\)/);
});

test("horizontal and vertical resize state are tracked separately", () => {
    assert.match(indexHtml, /data-canvas-resize-width-ready/);
    assert.match(indexHtml, /data-canvas-resize-height-ready/);
    assert.match(indexHtml, /canvasResizeWidthReady/);
    assert.match(indexHtml, /canvasResizeHeightReady/);

    // Horizontal locking must be guarded by a real right-edge movement so a
    // height-only drag does not freeze the panel's automatic width.
    assert.match(
        indexHtml,
        /if \(active\.right && Math\.abs\(deltaX\) >= 1\)[\s\S]*?canvasResizeWidthReady/s,
    );
    assert.match(
        indexHtml,
        /if \(active\.bottom && Math\.abs\(deltaY\) >= 1\)[\s\S]*?canvasResizeHeightReady/s,
    );
});

test("automatic width is rescheduled when visible controls change", () => {
    assert.match(indexHtml, /var scheduleAutoWidths = function \(\)/);
    assert.match(indexHtml, /var observer = new MutationObserver\(scheduleAutoWidths\)/);
    assert.match(indexHtml, /attributeFilter: \["class", "title"\]/);
});
