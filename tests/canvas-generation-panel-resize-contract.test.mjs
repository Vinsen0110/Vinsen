import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");

test("generation panels resize from the outer frame in both axes", () => {
    assert.match(
        indexHtml,
        /\.canvas-generation-panel \{[^}]*position: relative;[^}]*display: flex;[^}]*flex-direction: column;[^}]*min-width: 660px;[^}]*min-height: 0;[^}]*max-width: calc\(100vw - 32px\) !important;[^}]*resize: none !important;[^}]*overflow: visible !important;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel > \.relative\.h-full\.w-full \{[^}]*display: flex;[^}]*flex-direction: column;[^}]*flex: 1 1 auto;[^}]*width: 100%;[^}]*min-width: 0;[^}]*max-width: 100%;[^}]*min-height: 0;[^}]*height: auto !important;[^}]*box-sizing: border-box;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel > \.p-3 \{[^}]*flex: 1 1 auto;[^}]*width: 100%;[^}]*min-width: 0;[^}]*max-width: 100%;[^}]*box-sizing: border-box;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel > \.relative\.rounded-xl\.border \{[^}]*flex: 1 1 auto;[^}]*width: 100%;[^}]*min-width: 0;[^}]*max-width: 100%;[^}]*box-sizing: border-box;/s,
    );
    assert.match(indexHtml, /--canvas-panel-resize-width/);
    assert.match(indexHtml, /--canvas-panel-resize-height/);
    assert.match(indexHtml, /document\.addEventListener\("pointerdown", beginResize, true\)/);
    assert.match(indexHtml, /window\.addEventListener\("pointerup", finishResize, true\)/);
    assert.match(indexHtml, /var getMinimumWidth = function \(panel\)/);
    assert.match(indexHtml, /var updateAutoWidth = function \(panel\)/);
    assert.match(indexHtml, /panel\.dataset\.canvasWidthLocked === "true"/);
    assert.match(indexHtml, /var observer = new MutationObserver\(scheduleAutoWidths\)/);
    assert.match(indexHtml, /attributeFilter: \["class", "title"\]/);
    assert.match(indexHtml, /minWidth: panelMinWidth/);
    assert.match(indexHtml, /Math\.max\(active\.minWidth, Math\.min\(active\.maxWidth, active\.width \+ deltaX\)\)/);
    assert.match(indexHtml, /active\.panel\.dataset\.canvasWidthLocked = "true"/);
    assert.match(indexHtml, /if \(!active\.widthChanged\) scheduleAutoWidths\(\)/);
    assert.match(
        indexHtml,
        /\.canvas-generation-panel \{[^}]*width: min\(var\(--canvas-panel-auto-width, 660px\), calc\(100vw - 32px\)\) !important;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel\[data-canvas-width-locked="true"\] \{[^}]*width: var\(--canvas-panel-resize-width\) !important;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel\[data-canvas-resize-ready="true"\] \{[^}]*height: var\(--canvas-panel-resize-height\) !important;/s,
    );
});

test("prompt editors no longer expose an inner resize handle", () => {
    assert.match(
        indexHtml,
        /\.canvas-generation-panel\[data-canvas-resize-ready="true"\] textarea,[\s\S]*?\.canvas-generation-panel\[data-canvas-resize-ready="true"\] \[contenteditable="true"\] \{[^}]*flex: 1 1 auto;[^}]*height: auto !important;[^}]*max-height: none !important;[^}]*resize: none !important;/s,
    );
});

test("parameter controls keep fixed widths and stay on one row", () => {
    assert.match(
        indexHtml,
        /\.canvas-generation-toolbar \{[^}]*height: 48px !important;[^}]*min-height: 48px;[^}]*width: 100%;[^}]*max-width: 100%;[^}]*flex-wrap: nowrap;[^}]*align-content: center;[^}]*overflow: visible;[^}]*box-sizing: border-box;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-toolbar > div:first-child \{[^}]*flex: 1 1 auto;[^}]*flex-wrap: nowrap;[^}]*justify-content: space-between;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-toolbar > div:first-child \{[^}]*overflow: visible;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-panel\[data-canvas-width-locked="true"\] \.canvas-generation-toolbar > div:first-child \{[^}]*overflow-x: auto;[^}]*overflow-y: hidden;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-toolbar > div:last-child \{[^}]*margin-left: 4px !important;/s,
    );
    assert.match(
        indexHtml,
        /\.canvas-generation-toolbar \.canvas-model-select,[\s\S]*?\.canvas-generation-toolbar \.canvas-prompt-preset-select \{[^}]*flex: 0 0 auto;/s,
    );
});
