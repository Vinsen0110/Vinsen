import assert from "node:assert/strict";
import test from "node:test";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import {
    mentionMenuLayout, keepMentionVisible, observeMentionMenu,
    remapImageMentions, reorderReferenceEdges, observeFlowLine, installCanvasFlowActivity,
} from "../canvas-reference-ui.js";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const css = await readFile(new URL("../canvas-interaction.css", import.meta.url), "utf8");

test("mention candidates show a thumbnail and label without a decorative @ prefix", () => {
    const source = bundle.slice(bundle.indexOf("function Tze("), bundle.indexOf("function Pze("));
    assert.match(source, /children:\[y.jsx\(Pze,\{reference\}\),y.jsx\("span",\{children:reference.label\}\)\]/);
    assert.doesNotMatch(source, /children:"@"/);
    assert.match(bundle, /function lX\(e\)\{return`@\$\{e.label\}`\}/);
});

test("mention menu follows the anchor, clamps to viewport and retains its available side", () => {
    const bounds = { left: 8, top: 8, right: 792, bottom: 592 };
    const anchor = { left: 750, top: 500, bottom: 520 };
    assert.deepEqual(mentionMenuLayout(anchor, bounds, 240), {
        left: 536, top: 254, width: 256, height: 240, side: "top",
    });
    const moved = mentionMenuLayout({ ...anchor, top: 450, bottom: 470 }, bounds, 240, "top");
    assert.equal(moved.top, 204);
    const lower = mentionMenuLayout({ left: 20, top: 50, bottom: 70 }, bounds, 240, "top");
    assert.equal(lower.side, "bottom");
    assert.equal(lower.top, 76);
});

test("mention menu fits a narrow viewport and measures a short list without phantom height", () => {
    const bounds = { left: 8, top: 8, right: 208, bottom: 220 };
    const result = mentionMenuLayout({ left: -20, top: 50, bottom: 70 }, bounds, 72);
    assert.equal(result.left, 8);
    assert.equal(result.width, 200);
    assert.equal(result.height, 72);
    const cramped = mentionMenuLayout({ left: 40, top: 100, bottom: 120 }, bounds, 400);
    assert.ok(cramped.top >= bounds.top);
    assert.ok(cramped.top + cramped.height <= bounds.bottom);
});

test("keyboard selection scrolls only the menu, including wrap to first and last item", () => {
    const list = {
        children: Array.from({ length: 15 }, (_, i) => ({ offsetTop: i * 40, offsetHeight: 40 })),
        scrollTop: 0, clientHeight: 160,
    };
    keepMentionVisible(list, 8);
    assert.equal(list.scrollTop, 200);
    keepMentionVisible(list, 0);
    assert.equal(list.scrollTop, 0);
    keepMentionVisible(list, 14);
    assert.equal(list.scrollTop, 440);
    keepMentionVisible(list, 13);
    assert.equal(list.scrollTop, 440);
});

test("reorder preserves reference identity in both raw and @ prompt tokens without chained replacements", () => {
    assert.equal(remapImageMentions("@图片1 + 图片2 + @图片3", ["a", "b", "c"], ["c", "a", "b"]),
        "@图片2 + 图片3 + @图片1");
    const ids = Array.from({ length: 12 }, (_, i) => String(i));
    assert.equal(remapImageMentions("@图片12 @图片1", ids, [...ids].reverse()), "@图片1 @图片12");
});

test("removal clears only the removed token and renumbers surviving image references", () => {
    assert.equal(remapImageMentions("@图片1、@图片2、@图片3，@文本1", ["a", "b", "c"], ["a", "c"]),
        "@图片1、、@图片2，@文本1");
    assert.equal(remapImageMentions("@图片1", ["a"], []), "");
    assert.equal(remapImageMentions(undefined, [], []), undefined);
});

const edges = [
    { id: "a", fromNodeId: "a", toNodeId: "out" },
    { id: "text", fromNodeId: "text", toNodeId: "out" },
    { id: "b", fromNodeId: "b", toNodeId: "out" },
    { id: "unrelated", fromNodeId: "a", toNodeId: "other" },
    { id: "c", fromNodeId: "c", toNodeId: "out" },
];
test("reorder/remove changes only the targeted reference edges and never mutates the graph", () => {
    const original = structuredClone(edges);
    const reordered = reorderReferenceEdges(edges, "out", ["a", "b", "c"], ["c", "a"]);
    assert.deepEqual(reordered.map(edge => edge.id), ["c", "text", "a", "unrelated"]);
    assert.deepEqual(edges, original);
    assert.equal(reordered[0], edges[4]);
    assert.equal(reordered[3], edges[3]);
    assert.equal(reorderReferenceEdges(edges, "out", ["a", "b", "c"], ["a", "b", "c"]), edges);
    assert.equal(reorderReferenceEdges(edges, "out", ["a", "b", "c"], ["z"]), edges);
    assert.equal(reorderReferenceEdges(edges, "out", ["a", "b", "c"], ["a", "a"]), edges);
});

function callbackHarness(nodes, connections) {
    const context = {
        c: { useCallback: fn => fn }, Ne: { Image: "image", Config: "config" },
        vn: { current: nodes }, Jr: { current: connections },
        remapImageMentions, reorderReferenceEdges,
        _: value => { context.nodes = value; }, ee: value => { context.edges = value; },
    };
    const refs = bundle.slice(bundle.indexOf("function cze("), bundle.indexOf("function V6("));
    const callback = bundle.slice(bundle.indexOf("d1=c.useCallback"), bundle.indexOf("f1=c.useCallback")).trim().replace(/,$/, "");
    vm.runInNewContext(`${refs};var ${callback}`, context);
    return context;
}

test("actual panel callback updates live refs, preserves images and unrelated metadata, and supports a reversible snapshot", () => {
    const images = ["a", "b", "c"].map(id => ({
        id, type: "image", metadata: { content: `blob:${id}`, storageKey: `image:${id}` },
    }));
    const nodes = [...images, { id: "out", type: "image", metadata: {
        prompt: "@图片2 @图片3", model: "same-model", generationPanelWidth: 900,
    } }];
    const state = callbackHarness(nodes, edges);
    state.d1("out", ["c", "a"]);
    assert.equal(state.nodes[3].metadata.prompt, " @图片1");
    assert.equal(state.nodes[3].metadata.model, "same-model");
    assert.equal(state.nodes[3].metadata.generationPanelWidth, 900);
    assert.equal(state.vn.current, state.nodes);
    assert.equal(state.Jr.current, state.edges);
    for (let i = 0; i < 3; i++) assert.equal(state.nodes[i], nodes[i]);
    // History stores the original immutable nodes/edges together.
    assert.equal(nodes[3].metadata.prompt, "@图片2 @图片3");
    assert.equal(edges.length, 5);
});

test("shared configuration consumers keep their prompts attached to the same reference image", () => {
    const nodes = [
        ...["a", "b"].map(id => ({ id, type: "image", metadata: { content: id } })),
        { id: "out", type: "image", metadata: { prompt: "@图片1" } },
        { id: "other", type: "image", metadata: { prompt: "@图片2" } },
        { id: "config", type: "config", metadata: {} },
    ];
    const connections = ["a", "b", "out", "other"].map(id => ({ id, fromNodeId: id, toNodeId: "config" }));
    const state = callbackHarness(nodes, connections);
    state.d1("out", ["b", "a"]);
    assert.equal(state.nodes[2].metadata.prompt, "@图片2");
    assert.equal(state.nodes[3].metadata.prompt, "@图片1");
    assert.equal(state.edges.length, 4);
});

test("popup observer batches movement, clamps caret to editor and cleans all subscriptions", t => {
    const listeners = new Map(), frames = new Map(), observers = [];
    let next = 0, disconnected = 0;
    const events = {
        addEventListener: (key, fn) => listeners.set(key, fn),
        removeEventListener: key => listeners.delete(key),
    };
    const observer = class {
        constructor(fn) { this.callback = fn; observers.push(this); }
        observe() {}
        disconnect() { disconnected++; }
    };
    const globals = {
        window: { ...events, innerWidth: 800, innerHeight: 600 }, document: events,
        ResizeObserver: observer, MutationObserver: observer,
        requestAnimationFrame: fn => { frames.set(++next, fn); return next; },
        cancelAnimationFrame: id => frames.delete(id),
    };
    for (const [key, value] of Object.entries(globals)) {
        const previous = Object.getOwnPropertyDescriptor(globalThis, key);
        Object.defineProperty(globalThis, key, { value, configurable: true, writable: true });
        t.after(() => previous ? Object.defineProperty(globalThis, key, previous) : delete globalThis[key]);
    }
    let left = 20;
    const anchor = { isConnected: true, parentElement: null, closest: () => null,
        getBoundingClientRect: () => ({ left, right: left + 300, top: 40, bottom: 140 }) };
    const menu = { isConnected: true, style: {}, firstElementChild: { offsetHeight: 30 },
        querySelector: () => ({ scrollHeight: 400 }), contains: target => target === menu };
    const cleanup = observeMentionMenu(anchor, menu, () => ({ left: -100, top: -100, bottom: -50 }));
    assert.equal(menu.style.left, "20px");
    assert.equal(menu.style.top, "46px");
    left = 80;
    listeners.get("scroll")({});
    listeners.get("resize")({});
    observers[0].callback();
    assert.equal(frames.size, 1);
    const frame = [...frames.values()][0]; frames.clear(); frame();
    assert.equal(menu.style.left, "80px");
    assert.equal(frames.size, 0, "no perpetual animation-frame measurement loop");
    listeners.get("scroll")({ target: menu });
    assert.equal(frames.size, 0, "scrolling inside the list does not reposition it");
    cleanup();
    assert.equal(listeners.size, 0);
    assert.equal(disconnected, 2);
});

test("animated paths share one visibility observer and release it after unmount", t => {
    const instances = [];
    const original = globalThis.IntersectionObserver;
    globalThis.IntersectionObserver = class {
        constructor(callback) { this.callback = callback; this.lines = new Set(); instances.push(this); }
        observe(line) { this.lines.add(line); }
        unobserve(line) { this.lines.delete(line); }
        disconnect() { this.disconnected = true; }
    };
    t.after(() => original === undefined ? delete globalThis.IntersectionObserver : globalThis.IntersectionObserver = original);
    const lines = Array.from({ length: 24 }, () => ({ dataset: {} }));
    const cleanup = lines.map(observeFlowLine);
    assert.equal(instances.length, 1);
    assert.equal(lines[0].dataset.canvasFlowVisible, "false");
    instances[0].callback([{ target: lines[0], isIntersecting: true }]);
    assert.equal(lines[0].dataset.canvasFlowVisible, "true");
    cleanup.forEach(fn => fn());
    assert.equal(instances[0].lines.size, 0);
    assert.ok(instances[0].disconnected);
});

test("flow activity pauses canvas movement but not typing and cleans up event listeners", t => {
    const listeners = new Map();
    const events = { addEventListener: (key, fn) => listeners.set(key, fn), removeEventListener: key => listeners.delete(key) };
    let writes = 0;
    const attrs = new Set();
    const canvas = { hasAttribute: key => attrs.has(key), setAttribute: key => { attrs.add(key); writes++; }, removeAttribute: key => attrs.delete(key) };
    class Element {
        constructor(editor = false) { this.editor = editor; }
        closest(selector) { return selector === ".laowu-canvas-viewport" ? canvas : this.editor ? this : null; }
    }
    for (const [key, value] of Object.entries({ window: events, document: { ...events, documentElement: { dataset: {} }, hidden: false }, Element })) {
        const descriptor = Object.getOwnPropertyDescriptor(globalThis, key);
        Object.defineProperty(globalThis, key, { value, configurable: true, writable: true });
        t.after(() => descriptor ? Object.defineProperty(globalThis, key, descriptor) : delete globalThis[key]);
    }
    const cleanup = installCanvasFlowActivity();
    listeners.get("wheel")({ target: new Element(true) });
    assert.equal(writes, 0);
    for (let i = 0; i < 10; i++) listeners.get("pointermove")({ target: new Element(), buttons: 1 });
    assert.equal(writes, 1);
    listeners.get("pointerup")();
    assert.equal(attrs.size, 0);
    document.hidden = true;
    listeners.get("visibilitychange")();
    assert.equal(document.documentElement.dataset.canvasPageHidden, "true");
    cleanup();
    assert.equal(listeners.size, 0);
});

test("rendered lines use lightweight CSS, visibility pausing and reduced-motion preference", () => {
    const source = bundle.slice(bundle.indexOf("function MTe("), bundle.indexOf("function TTe("));
    assert.match(source, /observeFlowLine\(flow.current\)/);
    assert.match(source, /className:"canvas-connection-static"/);
    assert.doesNotMatch(source, /drop-shadow/);
    assert.match(css, /\.canvas-flow-line\s*\{[^}]*filter: none;[^}]*will-change: auto;/);
    assert.match(css, /data-canvas-flow-visible="false"/);
    assert.match(css, /animation-play-state: paused/);
    assert.match(css, /prefers-reduced-motion: reduce/);
});

test("editor blur saves the value without reopening mentions and caret restoration does not scroll the canvas", () => {
    const editor = bundle.slice(bundle.indexOf("const sX=c.forwardRef"), bundle.indexOf("function Tze("));
    const match = editor.match(/onBlur:(F=>\{[^}]+\})/);
    assert.ok(match);
    const calls = [];
    const element = {};
    const context = {
        x: { current: true }, v: { current: element },
        j: () => calls.push("close"), n: value => calls.push(value),
        XM: value => { assert.equal(value, element); return "saved prompt"; },
        m: { onBlur: () => calls.push("blur") },
    };
    vm.runInNewContext(`(${match[1]})({})`, context);
    assert.equal(context.x.current, false);
    assert.deepEqual(calls, ["close", "saved prompt", "blur"]);
    assert.match(bundle, /function K6\(e,t\)\{\s*e.focus\(\{preventScroll:!0\}\)/);
});

test("panel remaps its local prompt in the same event before the graph changes", () => {
    assert.ok(bundle.includes("handleReferenceOrder=(nodeId,ids)=>"));
    const refs = Array.from({ length: 12 }, (_, i) => ({ nodeId: String(i + 1) }));
    const order = refs.slice(1).map(item => item.nodeId);
    const calls = [];
    const context = {
        P: "@图片1 and @图片12", T: refs, remapImageMentions,
        $: value => calls.push(["prompt", value]),
        l: (id, ids) => calls.push(["graph", id, ids]),
    };
    // Parent and editor must never see new references with the old numbered prompt.
    const source = bundle.slice(bundle.indexOf("handleReferenceOrder=") + "handleReferenceOrder=".length).split("\n")[0].replace(/,$/, "");
    vm.runInNewContext(`(${source})("output",${JSON.stringify(order)})`, context);
    assert.deepEqual(calls[0], ["prompt", " and @图片11"]);
    assert.equal(calls[1][0], "graph");
    assert.equal(calls[1][1], "output");
});

test("thumbnail drag stays inside its strip and removal invokes reference ordering only", () => {
    const source = bundle.slice(bundle.indexOf("function dke("), bundle.indexOf("function fke("));
    const calls = [], jsx = (type, props) => ({ type, props });
    let hook = 0, stops = 0;
    const context = {
        c: { useState: () => [hook++ === 0 ? "a" : null, () => {}] },
        y: { jsx, jsxs: jsx }, no: (...args) => args.join(" "), of: "close-icon",
    };
    const component = vm.runInNewContext(`(${source})`, context);
    const tree = component({ nodeId: "out", references: ["a", "b", "c"].map(nodeId => ({ nodeId })),
        onOrderChange: (id, ids) => calls.push([id, Array.from(ids)]) });
    const event = { stopPropagation: () => stops++, preventDefault() {}, dataTransfer: {} };
    tree.props.children[2].props.children[0].props.onDrop(event);
    assert.deepEqual(calls[0], ["out", ["b", "c", "a"]]);
    tree.props.children[1].props.children[1].props.onClick(event);
    assert.deepEqual(calls[1], ["out", ["a", "c"]]);
    assert.equal(stops, 2);
    assert.doesNotMatch(source, /deleteStoredImages|removeAsset|onDeleteNode/);
});
