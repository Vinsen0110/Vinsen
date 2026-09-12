import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");
const panelScriptStart = indexHtml.indexOf('var panelSelector = ".canvas-generation-panel";');
assert.ok(panelScriptStart >= 0);
const panelScript = indexHtml.slice(panelScriptStart, indexHtml.indexOf("</script>", panelScriptStart));

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
    assert.match(panelScript, /var scheduleAutoWidths = function \(panel\)/);
    assert.match(panelScript, /var pendingPanels = new Set\(\)/);
    assert.match(panelScript, /if \(knownPanels\.has\(panel\)\) pendingPanels\.add\(panel\)/);
    assert.match(panelScript, /knownPanels\.forEach\(function \(_, element\) \{ pendingPanels\.add\(element\); \}\)/);
    assert.match(panelScript, /if \(autoWidthFrame !== null\) return/);
    assert.match(panelScript, /var panels = Array\.from\(pendingPanels\);\s*pendingPanels\.clear\(\)/);
    assert.match(panelScript, /if \(!element\.isConnected\) return;\s*syncPanelObservation\(element\);\s*updateAutoWidth\(element\)/);
    assert.match(panelScript, /window\.addEventListener\("resize", scheduleAutoWidths\)/);
    assert.match(panelScript, /document\.fonts\.ready\.then\(function \(\) \{ scheduleAutoWidths\(\); \}\)/);
});

test("toolbar changes stay local and body observation only discovers added or removed panels", () => {
    assert.match(panelScript, /entry\.toolbarObserver = new MutationObserver\(function \(\) \{ scheduleAutoWidths\(panel\); \}\)/);
    assert.match(
        panelScript,
        /entry\.toolbarObserver\.observe\(toolbar, \{\s*subtree: true,\s*childList: true,\s*characterData: true,\s*attributes: true,\s*attributeFilter: \["class", "title", "style", "hidden"\]\s*\}\)/,
    );
    assert.match(panelScript, /registerPanels\(document\.body\)/);
    assert.match(panelScript, /record\.addedNodes\.forEach\(registerPanels\)/);
    assert.match(panelScript, /root\.matches\(panelSelector\) \? \[root\] : Array\.from\(root\.querySelectorAll\(panelSelector\)\)/);
    assert.match(panelScript, /observer\.observe\(document\.body, \{\s*subtree: true,\s*childList: true\s*\}\)/);
    assert.ok(panelScript.includes('if (panel && !target.closest("textarea, [contenteditable=\\"true\\"]"))'));
    assert.doesNotMatch(panelScript, /new MutationObserver\(scheduleAutoWidths\)/);
    assert.doesNotMatch(panelScript, /document\.querySelectorAll\(panelSelector\)/);
});

test("one shared resize observer schedules only the owning panel", () => {
    assert.equal((panelScript.match(/new ResizeObserver\(/g) || []).length, 1);
    assert.match(panelScript, /var resizeOwners = new Map\(\)/);
    assert.match(panelScript, /var panel = resizeOwners\.get\(entry\.target\);\s*if \(panel\) scheduleAutoWidths\(panel\)/);
    assert.match(panelScript, /new Set\(\[panel, toolbar, controls, toolbar && toolbar\.lastElementChild\]\.filter\(Boolean\)\)/);
    assert.match(panelScript, /controls\.children\)\.forEach\(function \(element\) \{ elements\.add\(element\); \}\)/);
    assert.match(panelScript, /resizeOwners\.set\(element, panel\);\s*if \(resizeObserver\) resizeObserver\.observe\(element\)/);
    assert.match(panelScript, /if \(elements\.has\(element\)\) return;\s*if \(resizeObserver\) resizeObserver\.unobserve\(element\);\s*resizeOwners\.delete\(element\)/);
});

test("viewport observers respond to scale changes without treating translation as scale", () => {
    assert.match(panelScript, /var viewportObservers = new Map\(\)/);
    assert.match(panelScript, /new DOMMatrixReadOnly\(layer\.style\.transform\)/);
    assert.match(panelScript, /return \[matrix\.a, matrix\.b, matrix\.c, matrix\.d\]\.join\(","\)/);
    assert.match(panelScript, /panel\.closest\("\.laowu-canvas-viewport > \.origin-top-left"\)/);
    assert.match(panelScript, /if \(scale === viewportEntry\.scale\) return/);
    assert.match(panelScript, /viewportEntry\.panels\.forEach\(scheduleAutoWidths\)/);
    assert.match(panelScript, /viewportEntry\.observer\.observe\(layer, \{ attributes: true, attributeFilter: \["style"\] \}\)/);
});

test("removed panels release pending work and all owned observers", () => {
    const cleanupStart = panelScript.indexOf("var removeDetachedPanels = function ()");
    const cleanup = panelScript.slice(cleanupStart, panelScript.indexOf("var getPanel =", cleanupStart));
    assert.ok(cleanupStart >= 0);
    assert.match(cleanup, /if \(panel\.isConnected\) return/);
    assert.match(cleanup, /resizeObserver\.unobserve\(element\)/);
    assert.match(cleanup, /resizeOwners\.delete\(element\)/);
    assert.match(cleanup, /releaseViewport\(entry\.layer, panel\)/);
    assert.match(cleanup, /entry\.toolbarObserver\.disconnect\(\)/);
    assert.match(cleanup, /pendingPanels\.delete\(panel\)/);
    assert.match(cleanup, /knownPanels\.delete\(panel\)/);
    assert.match(panelScript, /if \(!entry\.panels\.size\) \{\s*entry\.observer\.disconnect\(\);\s*viewportObservers\.delete\(layer\)/);
    assert.match(panelScript, /if \(removed\) removeDetachedPanels\(\)/);
});
