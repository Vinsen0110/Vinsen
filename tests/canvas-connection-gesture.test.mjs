import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const ref = current => ({ current });
const plain = value => JSON.parse(JSON.stringify(value));

function between(startMarker, endMarker) {
    const start = bundle.indexOf(startMarker);
    const end = bundle.indexOf(endMarker, start);
    assert.ok(start >= 0 && end > start, `Missing shipped code: ${startMarker}`);
    return bundle.slice(start, end);
}

function callback(name, next, scope) {
    const source = between(`${name}=c.useCallback(`, `,${next}=c.useCallback(`).slice(name.length + 1);
    return vm.runInNewContext(`(${source})`, { c: { useCallback: fn => fn }, ...scope });
}

function mouse(options = {}) {
    const calls = { prevented: 0, stopped: 0 };
    return {
        calls,
        type: "pointerdown", pointerId: 7, isPrimary: true,
        button: 0, buttons: 1, clientX: 12, clientY: 12,
        currentTarget: { getBoundingClientRect: () => ({ left: 0, top: 0, width: 24, height: 24 }) },
        preventDefault() { calls.prevented++; },
        stopPropagation() { calls.stopped++; },
        ...options,
    };
}

function renderHandle(visible = true, scale = 1, side = "right") {
    const scope = {
        c: { useState: () => [false, () => {}] },
        y: { jsx: (type, props) => ({ type, props }) },
    };
    const render = vm.runInNewContext(
        `${between("function Y6(", "const hX=")};Y6`, scope,
    );
    const started = [];
    return { tree: render({ side, scale, visible, onMouseDown: event => started.push(event) }), started };
}

test("the transparent positioning wrapper absorbs edge presses without starting connections or canvas selection", () => {
    const { tree, started } = renderHandle();
    assert.equal(tree.props.style.pointerEvents, "auto");
    for (const type of ["pointerdown", "mousedown"]) {
        const event = mouse({ type, clientX: 1, clientY: 19 });
        tree.props[type === "pointerdown" ? "onPointerDown" : "onMouseDown"](event);
        assert.ok(event.calls.prevented > 0);
        assert.ok(event.calls.stopped > 0);
        assert.equal(started.length, 0);
    }
    assert.equal(tree.props.children.props.style.pointerEvents, "auto");
    assert.equal(tree.props.children.props.style.width, 24);
    assert.equal(tree.props.children.props.style.height, 24);
    assert.equal(typeof tree.props.children.props.onPointerDown, "function");
    assert.equal(renderHandle(false).tree.props.style.pointerEvents, "none");
    assert.equal(renderHandle(false).tree.props.children.props.style.pointerEvents, "none");
});

test("connection circle accepts primary center presses and rejects secondary buttons and square corners", () => {
    const { tree, started } = renderHandle();
    const down = tree.props.children.props.onPointerDown;
    const center = mouse();
    down(center);
    assert.equal(started.length, 1);
    assert.ok(center.calls.prevented > 0);
    assert.ok(center.calls.stopped > 0);
    down(mouse({ button: 2 }));
    down(mouse({ button: 1 }));
    down(mouse({ isPrimary: false }));
    down(mouse({ clientX: 1, clientY: 1 }));
    down(mouse({ clientX: 25, clientY: 12 }));
    assert.equal(started.length, 1);
});

test("connection wrapper absorbs clicks and double-clicks without invoking node previews or canvas actions", () => {
    const { tree, started } = renderHandle();
    const click = mouse({ type: "click" });
    tree.props.onClick(click);
    assert.ok(click.calls.stopped > 0);
    const doubleClick = mouse({ type: "dblclick" });
    tree.props.onDoubleClick(doubleClick);
    assert.ok(doubleClick.calls.prevented > 0);
    assert.ok(doubleClick.calls.stopped > 0);
    assert.equal(started.length, 0);
});

test("compatibility mousedown never starts another connection after pointerdown", () => {
    const { tree, started } = renderHandle();
    const event = mouse({ type: "mousedown" });
    tree.props.children.props.onMouseDown?.(event);
    assert.equal(started.length, 0);
});

test("connection circle retains a 24 screen-pixel diameter at different canvas zoom levels", () => {
    for (const scale of [0.1, 0.5, 1, 2]) {
        const { tree } = renderHandle(true, scale);
        assert.equal(tree.props.children.props.style.width * scale, 24);
        assert.equal(tree.props.children.props.style.height * scale, 24);
    }
});

test("revealing a connection handle never moves or scales its hit target during a fast select-then-drag", () => {
    for (const scale of [0.1, 0.5, 1, 2]) {
        for (const side of ["left", "right"]) {
            const visible = renderHandle(true, scale, side).tree;
            const hidden = renderHandle(false, scale, side).tree;
            for (const field of ["width", "height", "left", "right", "transform"]) {
                assert.equal(visible.props.style[field], hidden.props.style[field], `${side} at zoom ${scale}: ${field}`);
            }
            assert.equal(visible.props.style.transform, "translateY(-50%)");
            assert.doesNotMatch(visible.props.style.transition, /transform|translate|scale|\ball\b/);
            assert.doesNotMatch(hidden.props.style.transition, /transform|translate|scale|\ball\b/);
            assert.equal(visible.props.children.props.style.width, hidden.props.children.props.style.width);
            assert.equal(visible.props.children.props.style.height, hidden.props.children.props.style.height);
        }
    }
});

function gestureScope() {
    const state = { cursor: null, selection: "browser text", menu: null, connection: null, marquee: null };
    const scope = {
        state, Va: ref(null), yc: ref(null), xl: ref(null), ci: ref(null),
        connectionPointerRef: ref(null),
        i: ref({ getBoundingClientRect: () => ({ left: 0, top: 0, right: 1200, bottom: 900 }) }),
        T: ref({ isDraggingNode: false }), Ao: ref({ k: 1 }), S: ref(null),
        lastPointerWorldRef: ref(null),
        or: (x, y) => ({ x, y }),
        xt: value => { state.cursor = value; },
        Ce: value => { state.connection = value; },
        Ie: value => { state.target = value; },
        ve: value => { state.selectedConnection = value; },
        it: value => { state.marquee = value; },
        _e: value => { state.menu = value; },
        Nl: () => { state.dragEnded = true; },
        es: () => {},
        window: { getSelection: () => ({ removeAllRanges: () => { state.selection = ""; } }) },
        Ds: () => ({ nodeId: null, isNearNode: false }),
        dv: (from, to) => { state.completed = { from, to }; },
    };
    scope.releaseCanvasConnectionPointer = vm.runInNewContext(
        `${between("function releaseCanvasConnectionPointer(", "function Y6(")};releaseCanvasConnectionPointer`,
    );
    scope.Qo = callback("Qo", "cv", scope);
    scope.Sl = callback("Sl", "Ds", scope);
    return scope;
}

function begin(scope, options = {}) {
    const captured = new Set();
    const element = {
        setPointerCapture(id) { captured.add(id); },
        hasPointerCapture(id) { return captured.has(id); },
        releasePointerCapture(id) { captured.delete(id); },
    };
    const event = mouse({ currentTarget: element, ...options });
    callback("p1", "g1", scope)(event, "image-1", "source");
    return { event, element, captured };
}

test("starting a connection cancels conflicting gestures and clears native selection", () => {
    const scope = gestureScope();
    scope.ci.current = { existing: true };
    scope.yc.current = { existing: true };
    const { event, element, captured } = begin(scope);
    assert.ok(event.calls.prevented > 0);
    assert.ok(event.calls.stopped > 0);
    assert.equal(scope.state.selection, "");
    assert.equal(scope.state.dragEnded, true);
    assert.equal(scope.ci.current, null);
    assert.equal(scope.yc.current, null);
    assert.deepEqual(plain(scope.Va.current), { nodeId: "image-1", handleType: "source" });
    assert.deepEqual(plain(scope.state.cursor), { x: 12, y: 12 });
    assert.equal(scope.connectionPointerRef.current.element, element);
    assert.equal(scope.connectionPointerRef.current.pointerId, 7);
    assert.deepEqual(plain(scope.connectionPointerRef.current.connection), { nodeId: "image-1", handleType: "source" });
    assert.ok(captured.has(7));
});

test("secondary buttons, secondary pointers, and legacy mouse starts do not mutate gesture state", () => {
    for (const options of [{ button: 2 }, { isPrimary: false }, { pointerId: undefined }]) {
        const scope = gestureScope();
        begin(scope, options);
        assert.equal(scope.Va.current, null);
        assert.equal(scope.state.selection, "browser text");
        assert.equal(scope.state.dragEnded, undefined);
        assert.equal(scope.connectionPointerRef.current, null);
    }
});

test("cancel clears the connection and synchronous create-menu reference", () => {
    const scope = gestureScope();
    const { captured } = begin(scope);
    scope.yc.current = { existing: true };
    scope.Sl();
    assert.equal(scope.Va.current, null);
    assert.equal(scope.yc.current, null);
    assert.equal(scope.state.menu, null);
    assert.equal(scope.connectionPointerRef.current, null);
    assert.equal(captured.size, 0);
});

function release(scope) {
    return callback("Bu", "moveCanvasConnection", scope);
}

function move(scope) {
    const source = between("moveCanvasConnection=c.useCallback(", ";c.useEffect(()=>").slice("moveCanvasConnection=".length);
    return vm.runInNewContext(`(${source})`, { c: { useCallback: fn => fn }, ...scope, Bu: release(scope) });
}

test("release on blank canvas opens one create menu and synchronizes its ref before React renders", () => {
    const scope = gestureScope();
    const { captured } = begin(scope);
    const end = release(scope);
    end(mouse({ type: "pointerup", buttons: 0 }));
    assert.ok(scope.state.menu);
    assert.equal(scope.yc.current, scope.state.menu);
    const original = scope.state.menu;
    end(mouse({ type: "pointerup", buttons: 0, clientX: 200 }));
    assert.equal(scope.state.menu, original, "pointerup and mouseup must not open two menus");
    assert.equal(scope.connectionPointerRef.current, null);
    assert.equal(captured.size, 0);
});

test("immediate pointerup uses the gesture session snapshot even when the render connection ref is stale", () => {
    const scope = gestureScope();
    const { dispatch } = installListeners(scope);
    begin(scope);
    scope.Va.current = null;
    dispatch("pointerup", mouse({ type: "pointerup", buttons: 0, clientX: 700, clientY: 600 }));
    assert.ok(scope.state.menu, "a fast release must not depend on a rendered connection ref");
    assert.deepEqual(plain(scope.state.menu.connection), { nodeId: "image-1", handleType: "source" });
    assert.deepEqual(plain(scope.state.menu.position), { x: 700, y: 600 });
    assert.equal(scope.connectionPointerRef.current, null);
});

test("compatibility mouseup finishes an active session when pointerup is missing and never opens it twice", () => {
    const scope = gestureScope();
    const { dispatch } = installListeners(scope);
    begin(scope);
    dispatch("mouseup", mouse({ type: "mouseup", pointerId: undefined, buttons: 0, clientX: 700, clientY: 600 }));
    assert.ok(scope.state.menu, "mouseup is a fallback for an active pointer session");
    assert.deepEqual(plain(scope.state.menu.position), { x: 700, y: 600 });
    const opened = scope.state.menu;
    dispatch("mouseup", mouse({ type: "mouseup", pointerId: undefined, buttons: 0, clientX: 800 }));
    dispatch("pointerup", mouse({ type: "pointerup", buttons: 0, clientX: 800 }));
    assert.equal(scope.state.menu, opened);
    assert.equal(scope.connectionPointerRef.current, null);
});

test("unrelated or late pointercancel events cannot discard an active session or its pending create menu", () => {
    const scope = gestureScope();
    const { dispatch } = installListeners(scope);
    begin(scope);
    const session = scope.connectionPointerRef.current;
    dispatch("pointercancel", { pointerId: 99 });
    assert.equal(scope.connectionPointerRef.current, session);
    dispatch("pointerup", mouse({ type: "pointerup", buttons: 0, clientX: 700, clientY: 600 }));
    const opened = scope.state.menu;
    assert.ok(opened);
    dispatch("pointercancel", { pointerId: 7 });
    dispatch("pointercancel", { pointerId: 99 });
    assert.equal(scope.state.menu, opened);
    assert.equal(scope.yc.current, opened);
});

test("blur, right-click, and matching pointercancel clear the session even before the connection ref renders", () => {
    for (const type of ["blur", "contextmenu", "pointercancel"]) {
        const scope = gestureScope();
        const { dispatch } = installListeners(scope);
        begin(scope);
        scope.Va.current = null;
        dispatch(type, mouse({ type, button: type === "contextmenu" ? 2 : 0 }), type !== "blur");
        assert.equal(scope.connectionPointerRef.current, null, type);
        dispatch("mouseup", mouse({ type: "mouseup", pointerId: undefined, buttons: 0, clientX: 700, clientY: 600 }));
        assert.equal(scope.state.menu, null, `${type} must not leave an active session for a later mouseup`);
    }
});

test("release connects an eligible target or cancels near-node misses", () => {
    for (const hit of [{ nodeId: "image-2", isNearNode: true }, { nodeId: null, isNearNode: true }]) {
        const scope = gestureScope();
        begin(scope);
        scope.Ds = () => hit;
        release(scope)(mouse({ buttons: 0 }));
        assert.equal(scope.Va.current, null);
        assert.equal(scope.state.menu, null);
        assert.equal(scope.state.completed?.to, hit.nodeId || undefined);
    }
});

test("legacy mousemove with buttons zero cannot cancel a pointer connection or its create menu", () => {
    for (const hasMenu of [false, true]) {
        const scope = gestureScope();
        begin(scope);
        if (hasMenu) scope.yc.current = { connection: scope.Va.current };
        callback("es", "Bs", scope)(mouse({ type: "mousemove", pointerId: undefined, buttons: 0 }));
        assert.ok(scope.Va.current);
    }
});

test("queued zero-button pointermove cannot prematurely cancel or complete the gesture at its origin", () => {
    const scope = gestureScope();
    begin(scope);
    const connection = scope.Va.current;
    move(scope)(mouse({ type: "pointermove", button: -1, buttons: 0, clientX: 12, clientY: 12 }));
    assert.equal(scope.Va.current, connection);
    assert.equal(scope.state.menu, null);
    assert.equal(scope.connectionPointerRef.current.pointerId, 7);
    move(scope)(mouse({ type: "pointermove", button: -1, buttons: 1, clientX: 700, clientY: 600 }));
    release(scope)(mouse({ type: "pointerup", buttons: 0, clientX: 700, clientY: 600 }));
    assert.deepEqual(plain(scope.state.menu.position), { x: 700, y: 600 });
    assert.equal(scope.state.menu, scope.yc.current);
    assert.equal(scope.connectionPointerRef.current, null);
});

test("unrelated pointers cannot move, finish, or clear the active connection", () => {
    const scope = gestureScope();
    begin(scope);
    const connection = scope.Va.current;
    const cursor = scope.state.cursor;
    move(scope)(mouse({ type: "pointermove", pointerId: 99, clientX: 500 }));
    move(scope)(mouse({ type: "pointermove", pointerId: 99, button: -1, buttons: 0 }));
    release(scope)(mouse({ type: "pointerup", pointerId: 99, buttons: 0 }));
    assert.equal(scope.Va.current, connection);
    assert.equal(scope.state.cursor, cursor);
    assert.equal(scope.state.menu, null);
    assert.equal(scope.connectionPointerRef.current.pointerId, 7);
});

test("secondary-button releases cannot finish a primary-pointer connection", () => {
    const scope = gestureScope();
    begin(scope);
    release(scope)(mouse({ type: "pointerup", button: 2, buttons: 1 }));
    assert.equal(scope.state.menu, null);
    assert.equal(scope.connectionPointerRef.current.pointerId, 7);
});

test("pointer movement updates the hover target and cursor but preserves an already-open menu", () => {
    const scope = gestureScope();
    begin(scope);
    scope.Ds = () => ({ nodeId: "target-2", isNearNode: true });
    move(scope)(mouse({ type: "pointermove", button: -1, clientX: 500, clientY: 400 }));
    assert.equal(scope.xl.current, "target-2");
    assert.deepEqual(plain(scope.state.cursor), { x: 500, y: 400 });
    scope.yc.current = { existing: true };
    move(scope)(mouse({ type: "pointermove", button: -1, buttons: 0 }));
    assert.ok(scope.Va.current);
    assert.equal(scope.yc.current.existing, true);
});

test("releasing outside canvas bounds cancels the connection without creating an off-canvas menu", () => {
    const scope = gestureScope();
    begin(scope);
    release(scope)(mouse({ type: "pointerup", buttons: 0, clientX: 1300 }));
    assert.equal(scope.Va.current, null);
    assert.equal(scope.state.menu, null);
    assert.equal(scope.connectionPointerRef.current, null);
});

function installListeners(scope) {
    const listeners = new Map();
    const removed = [];
    let cleanup;
    scope.Bu = release(scope);
    scope.moveCanvasConnection = move(scope);
    scope.es = () => {};
    scope.Bs = () => {};
    scope.window.addEventListener = (name, fn, capture) => listeners.set(`${name}:${!!capture}`, { name, fn, capture });
    scope.window.removeEventListener = (name, fn, capture) => removed.push({ name, fn, capture });
    scope.document = {
        hidden: false,
        addEventListener: scope.window.addEventListener,
        removeEventListener: scope.window.removeEventListener,
    };
    const start = bundle.indexOf(";c.useEffect(()=>", bundle.indexOf("Bu=c.useCallback("));
    const end = bundle.indexOf(";const Lh=", start);
    assert.ok(start >= 0 && end > start);
    vm.runInNewContext(bundle.slice(start + 1, end), {
        ...scope, c: { useEffect: effect => { cleanup = effect(); } },
    });
    const dispatch = (name, event, capture = true) => {
        const listener = listeners.get(`${name}:${capture}`);
        if (name === "lostpointercapture" && !listener) return;
        assert.ok(listener, `Missing gesture listener: ${name}`);
        listener.fn(event);
    };
    return { dispatch, listeners, removed, cleanup };
}

test("clicking the old transparent edge then dragging the circle opens the create menu without a stale marquee", () => {
    const scope = gestureScope();
    const { dispatch } = installListeners(scope);
    const { tree, started } = renderHandle();
    let canvasPointerDown = 0;
    const edge = mouse({ clientX: 1, clientY: 19 });
    tree.props.onPointerDown(edge);
    if (!edge.calls.stopped) {
        canvasPointerDown++;
        scope.ci.current = { stale: true };
        scope.it(scope.ci.current);
    }
    tree.props.onMouseDown(mouse({ type: "mousedown", clientX: 1, clientY: 19 }));
    dispatch("pointerup", mouse({ type: "pointerup", buttons: 0 }));
    assert.equal(canvasPointerDown, 0);
    assert.equal(started.length, 0);
    assert.equal(scope.ci.current, null);
    begin(scope);
    const originalConnection = scope.Va.current;
    dispatch("lostpointercapture", { pointerId: 7, target: scope.i.current });
    assert.equal(scope.Va.current, originalConnection, "old viewport capture loss must not cancel the new circle connection");
    dispatch("pointermove", mouse({ type: "pointermove", button: -1, clientX: 700, clientY: 600 }));
    dispatch("pointerup", mouse({ type: "pointerup", buttons: 0, clientX: 700, clientY: 600 }));
    assert.deepEqual(plain(scope.state.menu.position), { x: 700, y: 600 });
    assert.equal(scope.ci.current, null);
    assert.equal(scope.state.marquee, null);
});

test("pointerup and compatibility mouseup clear stale marquee state without an active connection", () => {
    for (const type of ["pointerup", "mouseup"]) {
        const scope = gestureScope();
        const { dispatch } = installListeners(scope);
        scope.ci.current = { stale: true };
        scope.it(scope.ci.current);
        dispatch(type, mouse({ type, buttons: 0 }));
        assert.equal(scope.ci.current, null, type);
        assert.equal(scope.state.marquee, null, type);
        assert.equal(scope.state.menu, null, type);
    }
});

test("lost pointer capture preserves the session so the global pointerup still completes it", () => {
    const scope = gestureScope();
    const { dispatch } = installListeners(scope);
    const { element, captured } = begin(scope);
    const connection = scope.Va.current;
    dispatch("lostpointercapture", { pointerId: 7, target: scope.i.current });
    assert.equal(scope.Va.current, connection);
    dispatch("lostpointercapture", { pointerId: 7, target: element });
    assert.equal(scope.Va.current, connection, "queued event is stale while the element still has capture");
    captured.delete(7);
    dispatch("lostpointercapture", { pointerId: 99, target: element });
    assert.equal(scope.Va.current, connection);
    dispatch("lostpointercapture", { pointerId: 7, target: element });
    assert.equal(scope.Va.current, connection);
    assert.equal(scope.connectionPointerRef.current.pointerId, 7);
    dispatch("pointerup", mouse({ type: "pointerup", buttons: 0, clientX: 700, clientY: 600 }));
    assert.deepEqual(plain(scope.state.menu.position), { x: 700, y: 600 });
    assert.equal(scope.connectionPointerRef.current, null);
});

test("global capture listeners terminate connections and clear canceled or blurred gestures", () => {
    const scope = gestureScope();
    const { dispatch, listeners, removed, cleanup } = installListeners(scope);
    assert.equal(listeners.get("mouseup:true").capture, true);
    assert.equal(listeners.get("pointerup:true").capture, true);
    assert.equal(listeners.get("pointermove:true").fn, scope.moveCanvasConnection);
    for (const name of ["blur", "pointercancel"]) {
        const { element, captured } = begin(scope);
        scope.ci.current = { existing: true };
        scope.yc.current = { existing: true };
        if (name === "lostpointercapture") captured.delete(7);
        dispatch(name, { pointerId: 7, target: element }, name !== "blur");
        assert.equal(scope.Va.current, null, name);
        assert.equal(scope.ci.current, null, name);
        assert.equal(scope.yc.current, null, name);
        assert.equal(scope.state.dragEnded, true, name);
    }
    const { element } = begin(scope);
    element.releasePointerCapture = pointerId => {
        assert.equal(scope.connectionPointerRef.current, null, "clear the owner before lostpointercapture can fire");
        dispatch("lostpointercapture", { pointerId, target: element });
    };
    dispatch("pointercancel", { pointerId: 99 });
    dispatch("lostpointercapture", { pointerId: 99 });
    assert.equal(scope.connectionPointerRef.current.pointerId, 7);
    dispatch("pointerup", mouse({ type: "pointerup", buttons: 0 }));
    assert.ok(scope.yc.current, "pointerup must finish a connection even when mouseup is lost");
    const opened = scope.yc.current;
    dispatch("mouseup", mouse({ type: "mouseup", pointerId: undefined, buttons: 0 }));
    dispatch("lostpointercapture", { pointerId: 7 });
    assert.equal(scope.yc.current, opened, "compatibility mouseup and normal capture release preserve the open menu");
    const context = mouse({ button: 2 });
    dispatch("contextmenu", context);
    assert.equal(scope.Va.current, null);
    assert.ok(context.calls.prevented > 0);
    begin(scope);
    scope.document.hidden = true;
    dispatch("visibilitychange", {}, false);
    assert.equal(scope.Va.current, null);
    cleanup();
    for (const listener of listeners.values()) {
        assert.ok(removed.some(item => item.name === listener.name && item.fn === listener.fn && item.capture === listener.capture));
    }
});

test("decorative node titles and timers cannot become browser text selections", () => {
    const source = between("function Yze(", "function Zze(");
    assert.match(source, /function Yze[\s\S]*?select-none/);
    assert.match(source, /function Qze[\s\S]*?select-none/);
});
