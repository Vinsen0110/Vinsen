import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const html = await readFile(new URL("../index.html", import.meta.url), "utf8");
const scripts = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)].map(match => match[1]);
const panelScript = scripts.find(source => source.includes("var knownPanels = new Map()"));
const editorScript = scripts.find(source => source.includes('var editorSelector = "textarea, [contenteditable='));

function sourceFunction(script, name) {
    const declaration = script.indexOf(`var ${name} = function`);
    assert.ok(declaration >= 0, `Missing function ${name}`);
    const start = script.indexOf("function", declaration);
    let end = script.indexOf("};", start);
    for (; end >= 0; end = script.indexOf("};", end + 2)) {
        const source = script.slice(start, end + 1);
        try {
            new vm.Script(`(${source})`);
            return source;
        } catch {}
    }
    throw new Error(`Cannot parse function ${name}`);
}

class Element {
    constructor(id) {
        this.id = id;
        this.isConnected = true;
        this.dataset = {};
    }
}

function scheduler() {
    const panels = Array.from({ length: 30 }, (_, index) => new Element(`p${index}`));
    const frames = [];
    const measured = [];
    const synced = [];
    const context = vm.createContext({
        Element,
        knownPanels: new Map(panels.map(panel => [panel, {}])),
        pendingPanels: new Set(),
        autoWidthFrame: null,
        requestAnimationFrame: callback => { frames.push(callback); return frames.length; },
        syncPanelObservation: panel => synced.push(panel.id),
        updateAutoWidth: panel => measured.push(panel.id),
    });
    vm.runInContext(`var scheduleAutoWidths=${sourceFunction(panelScript, "scheduleAutoWidths")}`, context);
    return { panels, frames, measured, synced, context };
}

test("panel measurements coalesce per panel without scanning the other 29 panels", () => {
    const app = scheduler();
    app.context.scheduleAutoWidths(app.panels[17]);
    app.context.scheduleAutoWidths(app.panels[17]);
    assert.equal(app.frames.length, 1);
    app.frames.shift()();
    assert.deepEqual(app.measured, ["p17"]);
    assert.deepEqual(app.synced, ["p17"]);
    assert.equal(app.context.pendingPanels.size, 0);
});

test("unknown and detached panels cannot trigger a later measurement", () => {
    const app = scheduler();
    app.context.scheduleAutoWidths(new Element("unregistered"));
    assert.equal(app.frames.length, 0);
    app.context.scheduleAutoWidths(app.panels[4]);
    app.panels[4].isConnected = false;
    app.frames.shift()();
    assert.deepEqual(app.measured, []);
    assert.deepEqual(app.synced, []);
});

test("window events visit only the registered panel set", () => {
    const app = scheduler();
    app.context.scheduleAutoWidths({ type: "resize" });
    app.frames.shift()();
    assert.deepEqual(app.measured, app.panels.map(panel => panel.id));
});

test("panel discovery registers only panels within the supplied added subtree", () => {
    const panels = [new Element("new-a"), new Element("new-b")];
    const root = new Element("added-wrapper");
    const queries = [];
    root.matches = () => false;
    root.querySelectorAll = selector => { queries.push(selector); return panels; };
    const scheduled = [];
    const context = vm.createContext({
        Element,
        panelSelector: ".canvas-generation-panel",
        knownPanels: new Map(),
        syncPanelObservation: () => {},
        scheduleAutoWidths: panel => scheduled.push(panel.id),
    });
    const register = vm.runInContext(`(${sourceFunction(panelScript, "registerPanels")})`, context);
    register(root);
    assert.deepEqual(queries, [".canvas-generation-panel"]);
    assert.deepEqual(scheduled, ["new-a", "new-b"]);
    assert.equal(context.knownPanels.size, 2);
});

test("detached panel cleanup releases its resize, toolbar and zoom registrations", () => {
    const panel = new Element("removed");
    panel.isConnected = false;
    const live = new Element("live");
    const child = new Element("toolbar");
    const unobserved = [];
    const released = [];
    let disconnected = 0;
    const entry = {
        elements: new Set([panel, child]), layer: "world",
        toolbarObserver: { disconnect: () => { disconnected++; } },
    };
    const context = vm.createContext({
        knownPanels: new Map([[panel, entry], [live, { elements: new Set() }]]),
        pendingPanels: new Set([panel, live]),
        resizeOwners: new Map([[panel, panel], [child, panel]]),
        resizeObserver: { unobserve: element => unobserved.push(element.id) },
        releaseViewport: (layer, element) => released.push([layer, element.id]),
    });
    vm.runInContext(`(${sourceFunction(panelScript, "removeDetachedPanels")})()`, context);
    assert.deepEqual(unobserved, ["removed", "toolbar"]);
    assert.deepEqual(released, [["world", "removed"]]);
    assert.equal(disconnected, 1);
    assert.equal(context.knownPanels.has(panel), false);
    assert.equal(context.knownPanels.has(live), true);
    assert.equal(context.pendingPanels.has(panel), false);
    assert.equal(context.resizeOwners.size, 0);
});

test("zoom identity ignores translation while retaining scale changes", () => {
    const context = vm.createContext({
        DOMMatrixReadOnly: class { constructor(matrix) { Object.assign(this, matrix); } },
    });
    const scale = vm.runInContext(`(${sourceFunction(panelScript, "viewportScale")})`, context);
    const matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    const value = transform => scale({ style: { transform } });
    assert.equal(value(matrix), value({ ...matrix, e: 80, f: -40 }));
    assert.notEqual(value(matrix), value({ ...matrix, a: 2, d: 2 }));
});

test("manually locked text-panel width returns before any layout measurement", () => {
    const panel = new Element("locked");
    panel.dataset.canvasWidthLocked = "true";
    panel.querySelector = () => null;
    const context = vm.createContext({
        getMetrics: () => { throw new Error("A locked panel must not be measured"); },
    });
    const update = vm.runInContext(`(${sourceFunction(panelScript, "updateAutoWidth")})`, context);
    assert.doesNotThrow(() => update(panel));
});

function widthHarness({ savedWidth = 660, scale = 1, viewport = 1600 } = {}) {
    const panel = new Element("image");
    panel.dataset.canvasWidthLocked = "true";
    panel.querySelector = selector => selector === ".canvas-generation-toolbar .canvas-model-select" ? {} : null;
    const properties = new Map([
        ["--canvas-panel-resize-width", savedWidth + "px"],
        ["--canvas-panel-resize-height", "300px"],
    ]);
    let writes = 0;
    panel.style = {
        getPropertyValue: key => properties.get(key) || "",
        setProperty: (key, value) => { properties.set(key, value); writes++; },
    };
    let minimum = 660;
    const context = vm.createContext({
        window: { innerWidth: viewport }, minWidth: 660,
        getMinimumWidth: () => minimum,
        getMetrics: () => ({ scaleX: scale }),
    });
    const update = vm.runInContext(`(${sourceFunction(panelScript, "updateAutoWidth")})`, context);
    return {
        panel, properties, context,
        update: width => { minimum = width; update(panel); },
        writes: () => writes,
        renderedWidth: () => Math.max(
            parseFloat(properties.get("--canvas-panel-resize-width")),
            parseFloat(properties.get("--canvas-panel-auto-width")),
        ),
    };
}

test("saved Nano Banana panel grows for GPT controls without changing saved width or height", () => {
    const app = widthHarness();
    app.update(660);
    assert.equal(app.renderedWidth(), 660);
    app.update(1020);
    assert.equal(app.renderedWidth(), 1020);
    assert.equal(app.properties.get("--canvas-panel-resize-width"), "660px");
    assert.equal(app.properties.get("--canvas-panel-resize-height"), "300px");
    assert.equal(app.panel.dataset.canvasWidthLocked, "true");
    const writes = app.writes();
    app.update(1020);
    assert.equal(app.writes(), writes, "stable measurements do not start an observer write loop");
});

test("a wider manual size is preserved across model and channel changes", () => {
    const app = widthHarness({ savedWidth: 1250 });
    for (const width of [660, 800, 1020, 1100, 660]) {
        app.update(width);
        assert.equal(app.renderedWidth(), 1250);
    }
    assert.equal(app.properties.get("--canvas-panel-resize-height"), "300px");
});

test("reopened narrow saved panels remeasure current controls and auto width respects canvas zoom", () => {
    const reopened = widthHarness({ savedWidth: 700 });
    reopened.update(1020);
    assert.equal(reopened.renderedWidth(), 1020);
    const zoomed = widthHarness({ scale: 1.5, viewport: 1400 });
    zoomed.update(1200);
    assert.equal(zoomed.properties.get("--canvas-panel-auto-width"), "912px");
});

test("active resizing is not overridden by automatic layout", () => {
    const app = widthHarness();
    app.panel.dataset.canvasResizeActive = "true";
    app.context.getMetrics = () => { throw new Error("active resize must not be remeasured"); };
    app.update(1100);
    assert.equal(app.writes(), 0);
});

test("only image panels combine the saved width and current parameter minimum", () => {
    assert.match(html, /\.canvas-generation-panel\[data-canvas-width-locked="true"\]:has\(\.canvas-generation-toolbar \.canvas-model-select\)\s*\{[^}]*width:\s*max\(var\(--canvas-panel-resize-width\), var\(--canvas-panel-auto-width, 660px\)\) !important;/);
});

test("text panel width follows its controls while image panels retain their 660px floor", () => {
    const controls = { children: [{ offsetWidth: 190, scrollWidth: 190 }, { offsetWidth: 100, scrollWidth: 100 }] };
    const action = { offsetWidth: 84 };
    const parent = {};
    const toolbar = { firstElementChild: controls, lastElementChild: action, parentElement: parent };
    const panel = isText => ({
        querySelector: selector => selector === ".canvas-text-model-picker" ? isText : toolbar,
    });
    const context = vm.createContext({
        minWidth: 660,
        getComputedStyle: () => ({ columnGap: "8px", marginLeft: "0", marginRight: "0" }),
        getHorizontalExtras: element => element === toolbar ? 18 : 24,
    });
    const minimum = vm.runInContext(`(${sourceFunction(panelScript, "getMinimumWidth")})`, context);
    assert.equal(minimum(panel(true)), 458);
    assert.equal(minimum(panel(false)), 660);
    controls.children[0].offsetWidth = 550;
    assert.equal(minimum(panel(true)), 818, "longer model names must not overlap the action");
    controls.children[0].scrollWidth = 580;
    assert.equal(minimum(panel(false)), 848, "image controls include their gap, padding and unclipped label width");
    assert.match(html, /\.canvas-generation-panel:has\(\.canvas-text-model-picker\)\s*\{[^}]*min-width:\s*min\(420px, calc\(100vw - 32px\)\)/);
    assert.match(html, /\.canvas-generation-toolbar:has\(\.canvas-text-model-picker\) > div:first-child\s*\{[^}]*justify-content:\s*flex-start;[^}]*gap:\s*8px;/);
});

function editorHarness() {
    const queries = [];
    class EditorElement extends Element {
        constructor(id, { editor = false, inPanel = false, descendants = [] } = {}) {
            super(id);
            this.editor = editor;
            this.inPanel = inPanel;
            this.descendants = descendants;
            this.properties = new Map();
            this.style = {
                getPropertyValue: key => this.properties.get(key) || "",
                removeProperty: key => this.properties.delete(key),
            };
        }
        matches() { return this.editor; }
        closest() { return this.inPanel ? {} : null; }
        querySelectorAll(selector) { queries.push([this.id, selector]); return this.descendants; }
    }
    const root = new EditorElement("root");
    const observers = [];
    vm.runInNewContext(editorScript, {
        Element: EditorElement,
        document: { documentElement: root, body: root },
        MutationObserver: class {
            constructor(callback) { this.callback = callback; observers.push(this); }
            observe() {}
        },
    });
    queries.length = 0;
    return { EditorElement, queries, observer: observers[0] };
}

test("ancestor style changes do not normalize or scan editor subtrees", () => {
    const app = editorHarness();
    const world = new app.EditorElement("world");
    app.observer.callback([{ type: "attributes", target: world }]);
    assert.deepEqual(app.queries, []);
});

test("editor style normalization preserves unrelated styles and ignores outside editors", () => {
    const app = editorHarness();
    const inside = new app.EditorElement("inside", { editor: true, inPanel: true });
    const outside = new app.EditorElement("outside", { editor: true });
    const dimensions = ["width", "height", "min-width", "min-height", "max-width", "max-height", "resize"];
    for (const editor of [inside, outside]) {
        for (const key of dimensions) editor.properties.set(key, "100px");
        editor.properties.set("color", "red");
        app.observer.callback([{ type: "attributes", target: editor }]);
    }
    assert.equal(inside.dataset.canvasFluidEditor, "true");
    assert.equal(inside.properties.size, 1);
    assert.equal(inside.properties.get("color"), "red");
    assert.equal(outside.properties.size, dimensions.length + 1);
    assert.deepEqual(app.queries, []);
});

test("new editor discovery only scans the added subtree", () => {
    const app = editorHarness();
    const editor = new app.EditorElement("new-editor", { editor: true, inPanel: true });
    editor.properties.set("height", "240px");
    const wrapper = new app.EditorElement("new-wrapper", { descendants: [editor] });
    app.observer.callback([{ type: "childList", addedNodes: [wrapper] }]);
    assert.equal(editor.properties.has("height"), false);
    assert.deepEqual(app.queries, [["new-wrapper", 'textarea, [contenteditable="true"]']]);
});
