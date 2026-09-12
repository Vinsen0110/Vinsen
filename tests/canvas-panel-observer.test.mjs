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

test("manually locked width returns before any layout measurement", () => {
    const panel = new Element("locked");
    panel.dataset.canvasWidthLocked = "true";
    const context = vm.createContext({
        getMetrics: () => { throw new Error("A locked panel must not be measured"); },
    });
    const update = vm.runInContext(`(${sourceFunction(panelScript, "updateAutoWidth")})`, context);
    assert.doesNotThrow(() => update(panel));
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
