import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const start = bundle.indexOf("function connectionMenuLayout(");
const menuStart = bundle.indexOf("function zke(", start);
const end = bundle.indexOf("function eB(", menuStart);
assert.ok(start >= 0 && menuStart > start && end > menuStart, "The shipped bundle must contain the viewport-aware connection menu");
const layoutSource = bundle.slice(start, menuStart);
const componentSource = bundle.slice(menuStart, end);
const layout = vm.runInNewContext(`${layoutSource};connectionMenuLayout`);
const plain = value => JSON.parse(JSON.stringify(value));
const canvas = { left: 0, top: 80, right: 661, bottom: 760 };
const screen = { width: 1300, height: 800 };
const viewport = { x: 0, y: 0, k: 1 };
const measured = { width: 220, height: 164 };

test("a connection released near the right edge opens fully inside the canvas", () => {
    const result = layout(canvas, viewport, { x: 628, y: 370 }, measured, screen);
    assert.equal(result.left, 441);
    assert.equal(result.top, 450);
    assert.ok(result.left + measured.width <= canvas.right);
    assert.ok(result.top + measured.height <= canvas.bottom);
});

test("all four corners and off-canvas releases stay within the visible canvas/window intersection", () => {
    for (const point of [{ x: -100, y: -100 }, { x: 2000, y: -100 }, { x: -100, y: 2000 }, { x: 2000, y: 2000 }]) {
        const result = layout(canvas, viewport, point, measured, screen);
        assert.ok(result.left >= 8);
        assert.ok(result.top >= canvas.top);
        assert.ok(result.left + measured.width <= canvas.right);
        assert.ok(result.top + measured.height <= canvas.bottom);
    }
});

test("world coordinates and menu dimensions scale together when the viewport has enough room", () => {
    const spaciousCanvas = { left: 0, top: 80, right: 3000, bottom: 2200 };
    const spaciousScreen = { width: 3200, height: 2400 };
    for (const k of [0.1, 0.5, 1, 2, 5]) {
        const transform = { x: 40, y: 30, k };
        const point = { x: (180 - spaciousCanvas.left - transform.x) / k, y: (260 - spaciousCanvas.top - transform.y) / k };
        const scaled = { width: measured.width * k, height: measured.height * k };
        const result = layout(spaciousCanvas, transform, point, scaled, spaciousScreen);
        assert.equal(result.left, 180);
        assert.equal(result.top, 260);
        assert.equal(result.maxWidth, 220);
        assert.ok(result.left + scaled.width <= spaciousCanvas.right);
        assert.ok(result.top + scaled.height <= spaciousCanvas.bottom);
        assert.equal(Math.min(measured.width, result.maxWidth) * k, measured.width * k);
        assert.equal(Math.min(measured.height, result.maxHeight) * k, measured.height * k);
    }
});

test("scaled menu bounds stay inside every canvas corner and use local size limits at high zoom", () => {
    for (const k of [0.1, 0.5, 1, 2, 5]) {
        const transform = { x: 40, y: 30, k };
        const scaled = { width: measured.width * k, height: measured.height * k };
        for (const anchor of [{ x: 8, y: 80 }, { x: 661, y: 80 }, { x: 8, y: 760 }, { x: 661, y: 760 }]) {
            const point = { x: (anchor.x - canvas.left - transform.x) / k, y: (anchor.y - canvas.top - transform.y) / k };
            const result = layout(canvas, transform, point, scaled, screen);
            const screenWidth = Math.min(measured.width, result.maxWidth) * k;
            const screenHeight = Math.min(measured.height, result.maxHeight) * k;
            assert.equal(result.maxWidth, Math.min(220, (canvas.right - 8) / k));
            assert.equal(result.maxHeight, (canvas.bottom - canvas.top) / k);
            assert.ok(result.left >= 8, `left at zoom ${k}`);
            assert.ok(result.top >= canvas.top, `top at zoom ${k}`);
            assert.ok(result.left + screenWidth <= canvas.right + 1e-9, `right at zoom ${k}`);
            assert.ok(result.top + screenHeight <= canvas.bottom + 1e-9, `bottom at zoom ${k}`);
        }
    }
});

test("small viewports constrain menu width and scrollable height instead of hiding options", () => {
    const result = layout({ left: 20, top: 30, right: 200, bottom: 180 }, viewport, { x: 170, y: 130 }, measured, screen);
    assert.deepEqual(plain(result), { left: 20, top: 30, maxWidth: 180, maxHeight: 150 });
});

test("partially offscreen canvas respects window margins and missing canvas falls back to the window", () => {
    const result = layout({ left: -50, top: -20, right: 1700, bottom: 900 }, viewport, { x: 2000, y: 2000 }, measured, screen);
    assert.equal(result.left + measured.width, screen.width - 8);
    assert.equal(result.top + measured.height, screen.height - 8);
    const fallback = layout(null, viewport, { x: -500, y: -500 }, measured, screen);
    assert.equal(fallback.left, 8);
    assert.equal(fallback.top, 8);
});

test("the actual menu portals to body, measures itself, tracks resizing, and cleans up observers", () => {
    const listeners = new Map();
    const removed = [];
    const observed = [];
    let disconnected = false;
    let state;
    let cleanup;
    const body = {};
    const canvasElement = { getBoundingClientRect: () => canvas };
    const element = {
        style: {},
        getBoundingClientRect() {
            return { width: Math.min(220, parseFloat(this.style.maxWidth)), height: Math.min(164, parseFloat(this.style.maxHeight)) };
        },
    };
    const scope = {
        window: {
            innerWidth: screen.width, innerHeight: screen.height,
            addEventListener: (name, fn, capture) => listeners.set(name, { fn, capture }),
            removeEventListener: (name, fn, capture) => removed.push({ name, fn, capture }),
        },
        document: { body },
        ResizeObserver: class {
            constructor(fn) { this.fn = fn; }
            observe(target) { observed.push(target); }
            disconnect() { disconnected = true; }
        },
        c: {
            useRef: () => ({ current: element }),
            useState: initial => {
                state = initial;
                return [state, update => { state = typeof update === "function" ? update(state) : update; }];
            },
            useLayoutEffect: effect => { cleanup = effect(); },
        },
        xr: { light: { node: { stroke: "#ccc" } } }, fr: () => "light",
        y: { jsx: (type, props) => ({ type, props }), jsxs: (type, props) => ({ type, props }) },
        Lo: { createPortal: (child, target) => ({ child, target }) },
        eB: "menu-item", g7: "input-icon", Is: "image-icon",
    };
    const render = vm.runInNewContext(`${layoutSource};${componentSource};zke`, scope);
    const actions = [];
    const portal = render({
        pending: { position: { x: 628, y: 660 } },
        viewport, containerRef: { current: canvasElement },
        onCreate: action => actions.push(action), onClose: () => actions.push("close"),
    });
    assert.equal(portal.target, body);
    assert.match(portal.child.props.className, /^fixed /);
    assert.equal(portal.child.props.style.overflowY, "auto");
    assert.equal(portal.child.props.style.boxSizing, "border-box");
    assert.equal(portal.child.props.style.width, 220);
    assert.equal(portal.child.props.style.height, 164);
    assert.equal(state.left, 441);
    assert.equal(state.top, 596);
    assert.ok(observed.includes(canvasElement));
    assert.ok(observed.includes(element));
    const initial = state;
    listeners.get("resize").fn();
    assert.equal(state, initial, "unchanged measurements must not cause a render loop");
    scope.window.innerWidth = 400;
    listeners.get("resize").fn();
    assert.equal(state.left, 172);
    portal.child.props.children[1].props.onClick();
    portal.child.props.children[2].props.onClick();
    portal.child.props.children[3].props.onClick();
    assert.deepEqual(actions, ["image-input", "text-generation", "image-generation"]);
    cleanup();
    assert.equal(disconnected, true);
    for (const [name, listener] of listeners) {
        assert.ok(removed.some(item => item.name === name && item.fn === listener.fn && item.capture === listener.capture));
    }
});

test("the actual create menu keeps compact world dimensions and transforms proportionally with canvas zoom", () => {
    const scope = {
        window: { innerWidth: screen.width, innerHeight: screen.height },
        document: { body: {} },
        c: {
            useRef: () => ({ current: null }),
            useState: initial => [initial, () => {}],
            useLayoutEffect: () => {},
        },
        xr: { light: { node: { stroke: "#ccc" } } }, fr: () => "light",
        y: { jsx: (type, props) => ({ type, props }), jsxs: (type, props) => ({ type, props }) },
        Lo: { createPortal: (child, target) => ({ child, target }) },
        eB: "menu-item", g7: "input-icon", Is: "image-icon",
    };
    const render = vm.runInNewContext(`${layoutSource};${componentSource};zke`, scope);
    for (const k of [0.1, 0.5, 1, 2, 5]) {
        const { child } = render({
            pending: { position: { x: 200, y: 300 } },
            viewport: { x: 40, y: 30, k },
            containerRef: { current: null }, onCreate: () => {}, onClose: () => {},
        });
        assert.equal(child.props.style.width, 220, `width at zoom ${k}`);
        assert.equal(child.props.style.height, 164, `height at zoom ${k}`);
        assert.equal(child.props.style.padding, 8);
        assert.equal(child.props.style.transform, `scale(${k})`);
        assert.equal(child.props.style.transformOrigin, "top left");
        assert.equal(child.props.style.zoom, undefined);
        assert.match(child.props.className, /^fixed /);
    }
});

test("the canvas passes its real viewport and container to the menu", () => {
    assert.ok(bundle.includes("Ge?y.jsx(zke,{pending:Ge,viewport:oe,containerRef:i,onCreate:"));
});
