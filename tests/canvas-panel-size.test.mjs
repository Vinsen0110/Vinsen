import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { bindCanvasPanelSize, PANEL_RESIZE_EVENT } from "../canvas-panel-size.js";

function panel() {
    const element = new EventTarget();
    const properties = new Map();
    element.dataset = {};
    element.style = {
        setProperty: (key, value) => properties.set(key, value),
        removeProperty: key => properties.delete(key),
        getPropertyValue: key => properties.get(key) || "",
    };
    element.resize = detail => element.dispatchEvent(new CustomEvent(PANEL_RESIZE_EVENT, { detail }));
    return element;
}

test("manual dimensions survive deselection, remount and serialized metadata", () => {
    let metadata = { model: "apimart::gemini-3.8-flash", prompt: "original" };
    const first = panel();
    const cleanup = bindCanvasPanelSize(first, metadata, patch => { metadata = { ...metadata, ...patch }; });
    first.resize({ width: 810.5, height: 345 });
    cleanup();
    const reopened = panel();
    bindCanvasPanelSize(reopened, JSON.parse(JSON.stringify(metadata)), () => {});
    assert.equal(reopened.style.getPropertyValue("--canvas-panel-resize-width"), "810.5px");
    assert.equal(reopened.style.getPropertyValue("--canvas-panel-resize-height"), "345px");
    assert.equal(reopened.dataset.canvasWidthLocked, "true");
    assert.equal(reopened.dataset.canvasResizeHeightReady, "true");
    assert.equal(metadata.model, "apimart::gemini-3.8-flash");
    assert.equal(metadata.prompt, "original");
    first.resize({ width: 999 });
    assert.equal(metadata.generationPanelWidth, 810.5, "detached panels must release their listener");
});

test("nodes and manually adjusted axes stay independent", () => {
    const widthOnly = panel(), heightOnly = panel();
    const patches = [];
    bindCanvasPanelSize(widthOnly, { generationPanelWidth: 730 }, patch => patches.push(patch));
    bindCanvasPanelSize(heightOnly, { generationPanelHeight: 300 }, () => {});
    assert.equal(widthOnly.dataset.canvasResizeHeightReady, undefined);
    assert.equal(widthOnly.style.getPropertyValue("--canvas-panel-resize-height"), "");
    assert.equal(heightOnly.dataset.canvasWidthLocked, undefined);
    widthOnly.resize({ height: 280 });
    assert.deepEqual(patches, [{ generationPanelHeight: 280 }]);
    assert.equal(heightOnly.style.getPropertyValue("--canvas-panel-resize-height"), "300px");
});

test("invalid, unchanged and missing dimensions produce no metadata changes", () => {
    const element = panel(), patches = [];
    bindCanvasPanelSize(element, { generationPanelWidth: 800 }, patch => patches.push(patch));
    for (const detail of [{}, { width: 800 }, { width: NaN }, { width: Infinity }, { height: -3 }, { width: "900" }]) {
        element.resize(detail);
    }
    assert.deepEqual(patches, []);
});

test("restoring an old node or undoing a resize returns to automatic sizing", () => {
    const element = panel();
    const cleanup = bindCanvasPanelSize(element, { generationPanelWidth: 800, generationPanelHeight: 320 }, () => {});
    cleanup();
    bindCanvasPanelSize(element, { generationPanelWidth: -1, generationPanelHeight: null }, () => {});
    assert.deepEqual(element.dataset, {});
    assert.equal(element.style.getPropertyValue("--canvas-panel-resize-width"), "");
    assert.equal(element.style.getPropertyValue("--canvas-panel-resize-height"), "");
});

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const html = await readFile(new URL("../index.html", import.meta.url), "utf8");

test("normal and reverse generation roots bind the node-specific size ref", () => {
    assert.match(bundle, /const panelRef=useCanvasPanelSize\(e,r\)/);
    assert.match(bundle, /y.jsx\(lke,\{panelRef:panelRef,config:w/);
    assert.match(bundle, /ref:panelRef,className:"canvas-generation-panel rounded-2xl/);
    assert.match(bundle, /function lke\(\{panelRef:panelRef/);
    assert.match(bundle, /ref:panelRef,className:"canvas-generation-panel rounded-xl/);
    assert.match(bundle, /patch=>onConfigChange\(node.id,patch\)/);
});

test("pointer release saves only moved axes in unscaled CSS units", () => {
    const source = html.match(/var finishResize = (function \(event\) \{[\s\S]*?)\n                };/)[1] + "\n}";
    for (const [widthChanged, heightChanged] of [[false, false], [true, false], [false, true], [true, true]]) {
        const element = panel(), patches = [];
        bindCanvasPanelSize(element, {}, patch => patches.push(patch));
        element.style.setProperty("--canvas-panel-resize-width", "800.5px");
        element.style.setProperty("--canvas-panel-resize-height", "310px");
        if (heightChanged) element.dataset.canvasResizeHeightReady = "true";
        let scheduled = 0;
        const scope = vm.createContext({
            active: { panel: element, widthChanged, heightChanged },
            CustomEvent, scheduleAutoWidths(value) { assert.equal(value, element); scheduled++; },
        });
        vm.runInContext(`(${source})({pointerId:1})`, scope);
        assert.equal(scope.active, null);
        assert.equal(scheduled, 1, "reconcile required model width after every completed resize");
        const expected = {};
        if (widthChanged) expected.generationPanelWidth = 800.5;
        if (heightChanged) expected.generationPanelHeight = 310;
        assert.deepEqual(patches, Object.keys(expected).length ? [expected] : []);
        if (!heightChanged) assert.equal(element.style.getPropertyValue("--canvas-panel-resize-height"), "");
    }
});
