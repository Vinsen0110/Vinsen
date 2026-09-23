import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const layoutStart = bundle.indexOf("function connectionMenuLayout(");
const layoutEnd = bundle.indexOf("function zke(", layoutStart);
const menuStart = bundle.indexOf("function kke(");
const menuEnd = bundle.indexOf("function G0(", menuStart);
assert.ok(layoutStart >= 0 && layoutEnd > layoutStart && menuStart >= 0 && menuEnd > menuStart);
const source = `${bundle.slice(layoutStart, layoutEnd)};${bundle.slice(menuStart, menuEnd)};kke`;
const scales = [0.1, 0.5, 1, 2, 5];
const canvas = { left: 20, top: 80, right: 3000, bottom: 2200 };
const screen = { width: 3200, height: 2400 };

function renderMenu({ k = 1, rect = canvas, dimensions = screen, anchor = { x: 200, y: 300 } } = {}) {
    const actions = [];
    const listeners = new Map();
    const removed = [];
    let cleanup;
    class FakeElement {
        constructor(inside) { this.inside = inside; }
        closest(selector) {
            assert.equal(selector, "[data-canvas-quick-create]");
            return this.inside ? this : null;
        }
    }
    const scope = {
        Element: FakeElement,
        window: {
            innerWidth: dimensions.width, innerHeight: dimensions.height,
            addEventListener: (name, handler) => listeners.set(name, handler),
            removeEventListener: (name, handler) => removed.push({ name, handler }),
        },
        c: { useEffect: effect => { cleanup = effect(); } },
        y: { jsx: (type, props) => ({ type, props }), jsxs: (type, props) => ({ type, props }) },
        G0: "menu-action", g7: "input-icon", GP: "image-icon", cl: "video-icon",
    };
    const render = vm.runInNewContext(source, scope);
    const viewport = { x: 40, y: 30, k };
    const point = { x: (anchor.x - (rect?.left || 0) - viewport.x) / k, y: (anchor.y - (rect?.top || 0) - viewport.y) / k };
    const tree = render({
        menu: { x: 17, y: 19, position: point },
        viewport, containerRef: { current: rect ? { getBoundingClientRect: () => rect } : null },
        onCreate: action => actions.push(action), onClose: () => actions.push("close"),
    });
    return { tree, actions, listeners, removed, cleanup, FakeElement };
}

test("blank-canvas quick create uses fixed world dimensions and follows canvas zoom", () => {
    for (const k of scales) {
        const { tree } = renderMenu({ k });
        const { style } = tree.props;
        assert.equal(style.width, 160);
        assert.equal(style.height, 174);
        assert.equal(style.transform, `scale(${k})`);
        assert.equal(style.transformOrigin, "top left");
        assert.equal(style.boxSizing, "border-box");
        assert.match(tree.props.className, /^fixed /);
        assert.equal(style.zoom, undefined);
    }
});

test("quick create anchors to world position after pan and zoom, not stale click screen coordinates", () => {
    for (const k of scales) {
        const { tree } = renderMenu({ k });
        assert.equal(tree.props.style.left, 200);
        assert.equal(tree.props.style.top, 300);
        assert.ok(tree.props.style.left + 160 * k <= canvas.right);
        assert.ok(tree.props.style.top + 174 * k <= canvas.bottom);
    }
});

test("scaled quick-create menu stays inside all four viewport corners", () => {
    const rect = { left: 20, top: 80, right: 661, bottom: 760 };
    const dimensions = { width: 1300, height: 800 };
    for (const k of scales) {
        for (const anchor of [{ x: 20, y: 80 }, { x: 661, y: 80 }, { x: 20, y: 760 }, { x: 661, y: 760 }]) {
            const { tree } = renderMenu({ k, rect, dimensions, anchor });
            const { style } = tree.props;
            const width = Math.min(160, style.maxWidth) * k;
            const height = Math.min(174, style.maxHeight) * k;
            assert.ok(style.left >= rect.left, `left at zoom ${k}`);
            assert.ok(style.top >= rect.top, `top at zoom ${k}`);
            assert.ok(style.left + width <= rect.right + 1e-9, `right at zoom ${k}`);
            assert.ok(style.top + height <= rect.bottom + 1e-9, `bottom at zoom ${k}`);
        }
    }
});

test("small canvases constrain local menu bounds and keep overflow scrollable", () => {
    const rect = { left: 20, top: 30, right: 200, bottom: 180 };
    const { tree } = renderMenu({ k: 5, rect, anchor: { x: 200, y: 180 } });
    assert.equal(tree.props.style.left, 20);
    assert.equal(tree.props.style.top, 30);
    assert.equal(tree.props.style.maxWidth, 36);
    assert.equal(tree.props.style.maxHeight, 30);
    assert.equal(tree.props.style.overflowY, "auto");
});

test("all five quick-create actions and canvas event isolation are preserved", () => {
    const { tree, actions } = renderMenu();
    assert.equal(tree.props.children.length, 5);
    for (const child of tree.props.children) child.props.onClick();
    assert.deepEqual(actions, ["image-input", "image-generation", "video-input", "text-generation", "video-generation"]);
    let stopped = 0;
    tree.props.onPointerDown({ stopPropagation: () => { stopped++; } });
    tree.props.onMouseDown({ stopPropagation: () => { stopped++; } });
    tree.props.onWheel({ stopPropagation: () => { stopped++; } });
    assert.equal(stopped, 3);
    assert.equal(tree.props["data-canvas-quick-create"], true);
    assert.equal(tree.props["data-canvas-no-zoom"], true);
});

test("outside click closes quick create while inside clicks stay open and listener cleanup is exact", () => {
    const { actions, listeners, removed, cleanup, FakeElement } = renderMenu();
    const handler = listeners.get("pointerdown");
    assert.equal(typeof handler, "function");
    handler({ target: new FakeElement(true) });
    assert.deepEqual(actions, []);
    handler({ target: new FakeElement(false) });
    assert.deepEqual(actions, ["close"]);
    handler({ target: {} });
    assert.deepEqual(actions, ["close", "close"]);
    cleanup();
    assert.ok(removed.some(item => item.name === "pointerdown" && item.handler === handler));
});

test("the actual canvas supplies its viewport and container to blank-canvas quick create", () => {
    const invocation = bundle.match(/y\.jsx\(kke,\s*\{[^]*?\}\):null/)?.[0];
    assert.ok(invocation);
    assert.match(invocation, /viewport:oe/);
    assert.match(invocation, /containerRef:i/);
});
