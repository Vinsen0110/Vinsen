import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { moveCanvasGroupFrames, updateCanvasGroup, remapClonedCanvasGroups, finishCanvasGroupCloneDrag } from "../canvas-groups.js";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const ref = current => ({ current });
function callback(name, next, scope) {
    const start = bundle.indexOf(`${name}=c.useCallback(`) + name.length + 1;
    const end = bundle.indexOf(next ? `,${next}=c.useCallback(` : ";c.useEffect(()=>", start);
    assert.ok(start > name.length && end > start);
    return vm.runInNewContext(`(${bundle.slice(start, end)})`, scope);
}
function harness(scale = 1) {
    const state = {
        nodes: ["a", "b", "other"].map((id, index) => ({
            id, type: "image", title: id, position: { x: index * 300, y: 100 },
            width: 200, height: 150, metadata: { model: "gpt-image-2.5", content: "blob:original" },
            ...(index < 2 ? { canvasGroup: { id: "g", name: "Group", color: "#22a06b" } } : {}),
        })),
        selected: new Set(["a", "b"]),
    };
    const listeners = new Map(), frames = new Map();
    let frameId = 0;
    const scope = {
        moveCanvasGroupFrames, remapClonedCanvasGroups, finishCanvasGroupCloneDrag,
        c: { useCallback: fn => fn, useEffect: fn => { state.cleanup = fn(); } },
        vn: ref(state.nodes), Jr: ref([]), oo: ref(state.selected), Ao: ref({ k: scale }),
        T: ref({ isDraggingNode: false }), x: ref(false), E: ref(false), Ph: ref([]), S: ref(null),
        ci: ref(null), Va: ref(null), connectionPointerRef: ref(null),
        Ne: { Text: "text", Image: "image" }, ua: "loading", rd: "idle", lastPointerWorldRef: ref(null),
        _: update => { state.nodes = typeof update === "function" ? update(state.nodes) : update; scope.vn.current = state.nodes; },
        ye: ids => { state.selected = ids; scope.oo.current = ids; },
        ee() {}, ut() {}, He() {}, un() {}, ve() {}, ln() {}, Ii() {}, it() {}, Sl() {}, Bu() {},
        Ds() {}, Ie() {}, xt() {}, xl: ref(null), Bs() {},
        or: (x, y) => ({ x: x / scale, y: y / scale }),
        requestAnimationFrame: fn => { const id = ++frameId; frames.set(id, fn); return id; },
        cancelAnimationFrame: id => frames.delete(id),
        window: {
            addEventListener: (type, fn, capture) => listeners.set(`${type}:${!!capture}`, fn),
            removeEventListener: (type, fn, capture) => listeners.delete(`${type}:${!!capture}`),
        },
        document: { hidden: false, addEventListener() {}, removeEventListener() {} },
    };
    const sanitizeStart = bundle.indexOf("function sanitizeCanvasNodeClone(");
    const sanitizeEnd = bundle.indexOf("function zc(", sanitizeStart);
    scope.sanitizeCanvasNodeClone = vm.runInNewContext(
        `${bundle.slice(sanitizeStart, sanitizeEnd)};sanitizeCanvasNodeClone`, scope);
    scope.Dh = callback("Dh", "Nl", scope);
    scope.Nl = callback("Nl", "es", scope);
    scope.es = callback("es", "Bs", scope);
    scope.moveCanvasConnection = callback("moveCanvasConnection", null, scope);
    const effectStart = bundle.indexOf(";c.useEffect(()=>", bundle.indexOf("moveCanvasConnection=c.useCallback("));
    const effectEnd = bundle.indexOf(";const Lh=", effectStart);
    vm.runInNewContext(bundle.slice(effectStart + 1, effectEnd), scope);
    const groupStart = bundle.match(/startDrag:(\(event,id\)=>\{Dh\(event,id\);T\.current\.groupPointerId=event\.pointerId\})/);
    assert.ok(groupStart);
    const start = vm.runInNewContext(`(${groupStart[1]})`, scope);
    const event = overrides => ({
        pointerId: 7, button: 0, buttons: 1, clientX: 100, clientY: 200,
        preventDefault() {}, stopPropagation() {}, ...overrides,
    });
    return {
        scope, state, event,
        start: () => start(event(), "a"),
        dispatch: (type, data) => listeners.get(`${type}:true`)(event({ type, ...data })),
        flush: () => { for (const [id, fn] of frames) { frames.delete(id); fn(); } },
    };
}

for (const scale of [0.5, 1, 2]) {
    test(`actual group pointer-only dragging updates members and release at zoom ${scale}`, () => {
        const app = harness(scale);
        const original = app.state.nodes;
        app.start();
        app.dispatch("pointermove", { clientX: 160, clientY: 240 });
        app.flush();
        assert.equal(app.state.nodes[0].position.x, 60 / scale);
        assert.equal(app.state.nodes[1].position.y, 100 + 40 / scale);
        assert.equal(app.state.nodes[2], original[2]);
        app.dispatch("pointerup", { buttons: 0, clientX: 180, clientY: 260 });
        assert.equal(app.state.nodes[0].position.x, 80 / scale);
        assert.equal(app.state.nodes[1].position.y, 100 + 60 / scale);
        assert.equal(app.state.nodes[0].metadata, original[0].metadata);
        assert.equal(app.state.nodes[0].canvasGroup, original[0].canvasGroup);
        assert.equal(app.scope.T.current.isDraggingNode, false);
        assert.equal(app.scope.x.current, false);
        app.state.cleanup();
    });
}

for (const scale of [0.5, 1, 2]) {
    test(`actual Alt-copy release inherits original group only for inside drops at zoom ${scale}`, () => {
        for (const dx of [40, 900]) {
            const app = harness(scale);
            app.state.nodes = updateCanvasGroup(app.state.nodes, "g", {
                frame: { x: -24, y: 20, width: 600, height: 400 },
            });
            app.scope.vn.current = app.state.nodes;
            app.scope.oo.current = new Set(["a"]);
            const before = app.state.nodes;
            app.scope.Dh(app.event({ altKey: true }), "a");
            const copyId = [...app.state.selected][0];
            app.scope.es(app.event({ clientX: 100 + dx * scale, clientY: 220 }));
            app.flush();
            app.dispatch("pointerup", { buttons: 0, clientX: 100 + dx * scale, clientY: 220 });
            const copy = app.state.nodes.find(node => node.id === copyId);
            assert.equal(copy.canvasGroup?.id, dx === 40 ? "g" : undefined);
            assert.equal(copy.position.x, dx);
            assert.equal(copy.metadata.model, "gpt-image-2.5");
            assert.equal(copy.metadata.followActiveImageSite, true);
            assert.equal(app.state.nodes[0], before[0]);
            assert.equal(app.state.nodes[2], before[2]);
            app.state.cleanup();
        }
    });
}

test("actual Alt-click without dragging joins its source group on release", () => {
    const app = harness();
    app.scope.oo.current = new Set(["a"]);
    app.scope.Dh(app.event({ altKey: true }), "a");
    const copyId = [...app.state.selected][0];
    app.dispatch("pointerup", { buttons: 0 });
    assert.equal(app.state.nodes.find(node => node.id === copyId).canvasGroup.id, "g");
    app.state.cleanup();
});

test("group release commits its final coordinates even without an intermediate move event", () => {
    const app = harness();
    app.start();
    app.dispatch("pointerup", { buttons: 0, clientX: 190, clientY: 250 });
    assert.equal(app.state.nodes[0].position.x, 90);
    assert.equal(app.state.nodes[1].position.x, 390);
    assert.equal(app.scope.T.current.isDraggingNode, false);
});

test("foreign pointers and stale zero-button moves cannot move or cancel a group drag", () => {
    const app = harness();
    app.start();
    app.dispatch("pointermove", { pointerId: 8, clientX: 900 });
    app.dispatch("pointermove", { buttons: 0, clientX: 900 });
    app.dispatch("pointerup", { pointerId: 8, buttons: 0, clientX: 900 });
    app.dispatch("pointercancel", { pointerId: 8 });
    app.flush();
    assert.equal(app.scope.T.current.isDraggingNode, true);
    assert.equal(app.state.nodes[0].position.x, 0);
    app.dispatch("pointercancel", { pointerId: 7 });
    assert.equal(app.scope.T.current.isDraggingNode, false);
});

for (const scale of [0.5, 1, 2]) {
    test(`manual group frame follows actual pointer drag and final release at zoom ${scale}`, () => {
        const app = harness(scale);
        const frame = { x: -40, y: 20, width: 750, height: 380 };
        app.state.nodes = updateCanvasGroup(app.state.nodes, "g", { frame });
        app.scope.vn.current = app.state.nodes;
        const original = app.state.nodes;
        app.start();
        app.dispatch("pointermove", { clientX: 160, clientY: 240 });
        app.flush();
        assert.equal(app.state.nodes[0].canvasGroup.frame.x, frame.x + 60 / scale);
        app.dispatch("pointerup", { buttons: 0, clientX: 180, clientY: 260 });
        assert.deepEqual(app.state.nodes[0].canvasGroup.frame, {
            ...frame, x: frame.x + 80 / scale, y: frame.y + 60 / scale,
        });
        assert.deepEqual(app.state.nodes[1].canvasGroup.frame, app.state.nodes[0].canvasGroup.frame);
        assert.equal(app.state.nodes[2], original[2]);
        assert.equal(app.state.nodes[0].metadata, original[0].metadata);
        app.state.cleanup();
    });
}
