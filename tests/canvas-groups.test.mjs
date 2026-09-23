import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import {
    canvasGroupOf, collectCanvasGroups, canvasGroupBounds, assignCanvasGroup,
    createCanvasGroup, updateCanvasGroup, dissolveCanvasGroup,
    remapClonedCanvasGroups, placeCanvasGroupClones, isCanvasGroupShortcut, useCanvasGroups, createCanvasGroupComponents,
} from "../canvas-groups.js";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const css = await readFile(new URL("../canvas-groups.css", import.meta.url), "utf8");
const node = (id, x = 0, y = 0) => ({
    id, type: "image", title: id, position: { x, y }, width: 200, height: 100,
    metadata: { model: "gpt-image-2.5", prompt: "@\u56fe\u72471", generationPanelWidth: 1250,
        content: "blob:original", storageKey: "original-key", status: "success" },
});
const fixture = () => [node("a", -100, 20), node("b", 240, 190), node("c", 600, 400)];
const grouped = () => createCanvasGroup(fixture(), ["a", "b"], () => "group-1");
const groupId = nodes => nodes[0].canvasGroup.id;
const plain = value => JSON.parse(JSON.stringify(value));

test("grouping preserves positions, sizes, all provider metadata, and unrelated node identity", () => {
    const nodes = fixture();
    const result = createCanvasGroup(nodes, ["a", "b"], () => "group-1");
    assert.equal(result[0].canvasGroup.name, "\u5206\u7ec4 1");
    assert.equal(result[1].canvasGroup.id, "group-1");
    for (let i = 0; i < nodes.length; i++) {
        assert.equal(result[i].position, nodes[i].position);
        assert.equal(result[i].metadata, nodes[i].metadata);
        assert.equal(result[i].width, nodes[i].width);
        assert.equal(result[i].height, nodes[i].height);
        assert.equal(nodes[i].canvasGroup, undefined);
    }
    assert.equal(result[2], nodes[2]);
});

test("group creation requires two actual selected nodes and skips existing default names", () => {
    const nodes = fixture();
    assert.equal(createCanvasGroup(nodes, ["a", "missing"]), nodes);
    const first = grouped();
    const second = createCanvasGroup(first, ["b", "c"], () => "group-2");
    assert.equal(second[1].canvasGroup.name, "\u5206\u7ec4 2");
    assert.equal(collectCanvasGroups(second).length, 2);
    assert.equal(second[1].canvasGroup.id, second[2].canvasGroup.id);
    assert.equal(second[0].canvasGroup.id, "group-1");
});

test("renaming and coloring update every member without altering other groups", () => {
    const nodes = grouped();
    const renamed = updateCanvasGroup(nodes, "group-1", { name: "  References  ", color: "#4285e8" });
    assert.equal(renamed[0].canvasGroup.name, "References");
    assert.deepEqual(renamed[0].canvasGroup, renamed[1].canvasGroup);
    assert.equal(renamed[2], nodes[2]);
    assert.equal(nodes[0].canvasGroup.name, "\u5206\u7ec4 1");
    assert.equal(updateCanvasGroup(renamed, "group-1", { name: "References" }), renamed);
    assert.equal(updateCanvasGroup(renamed, "missing", {}), renamed);
});

test("joining replaces membership instead of nesting; removing and dissolving retain images", () => {
    const nodes = grouped();
    const joined = assignCanvasGroup(nodes, ["c"], nodes[0].canvasGroup);
    assert.equal(collectCanvasGroups(joined)[0].nodes.length, 3);
    const removed = assignCanvasGroup(joined, ["a"], null);
    assert.equal(removed[0].canvasGroup, undefined);
    const dissolved = dissolveCanvasGroup(removed, "group-1");
    assert.equal(collectCanvasGroups(dissolved).length, 0);
    assert.equal(dissolved.length, 3);
    dissolved.forEach((item, i) => assert.equal(item.metadata, nodes[i].metadata));
    assert.equal(dissolveCanvasGroup(dissolved, "missing"), dissolved);
});

test("project JSON roundtrip preserves groups, old projects remain ungrouped", () => {
    const nodes = grouped();
    const project = { version: 1, nodes, connections: [{ id: "edge", fromNodeId: "a", toNodeId: "b" }] };
    assert.deepEqual(collectCanvasGroups(plain(project).nodes), collectCanvasGroups(nodes));
    assert.deepEqual(collectCanvasGroups(fixture()), []);
    assert.equal(canvasGroupOf({ canvasGroup: { name: "broken" } }), null);
    assert.equal(canvasGroupOf({ canvasGroup: { id: "g", name: "", color: "url(bad)" } }).color, "#22a06b");
});

test("bounds adapt to negative positions, member moves and manual sizes without moving nodes", () => {
    const nodes = grouped().slice(0, 2);
    assert.deepEqual(canvasGroupBounds(nodes), { x: -124, y: -44, width: 588, height: 358 });
    const moved = nodes.map(item => ({ ...item, position: { x: item.position.x + 42, y: item.position.y - 30 } }));
    assert.deepEqual(canvasGroupBounds(moved), { x: -82, y: -74, width: 588, height: 358 });
    const resized = [nodes[0], { ...nodes[1], width: 1250 }];
    assert.equal(canvasGroupBounds(resized).width, 1638);
    assert.equal(canvasGroupBounds([]), null);
});

test("group copies get independent ids and partial single-node copies have no inherited group", () => {
    const nodes = grouped();
    const cloned = remapClonedCanvasGroups(nodes, () => "new-group");
    assert.equal(groupId(cloned), "new-group");
    assert.equal(cloned[1].canvasGroup.id, "new-group");
    assert.equal(groupId(nodes), "group-1");
    assert.equal(cloned[2], nodes[2]);
    assert.equal(remapClonedCanvasGroups([nodes[0]])[0].canvasGroup, undefined);
    const recolored = updateCanvasGroup(cloned, "new-group", { color: "#d75a87" });
    assert.equal(nodes[0].canvasGroup.color, "#22a06b");
    assert.equal(recolored[0].canvasGroup.color, "#d75a87");
});

test("Option-G uses physical KeyG and does not hijack composition, repeat or other shortcuts", () => {
    const event = { code: "KeyG", key: "\u00a9", altKey: true };
    assert.equal(isCanvasGroupShortcut(event), true);
    for (const patch of [{ ctrlKey: true }, { metaKey: true }, { shiftKey: true }, { repeat: true },
        { isComposing: true }, { defaultPrevented: true }, { code: "KeyH" }, { altKey: false }])
        assert.equal(isCanvasGroupShortcut({ ...event, ...patch }), false);
});

function hookHarness() {
    const listeners = new Map(), cleanups = [];
    const oldWindow = globalThis.window;
    globalThis.window = {
        addEventListener: (name, fn) => listeners.set(name, fn),
        removeEventListener: (name, fn) => { if (listeners.get(name) === fn) listeners.delete(name); },
    };
    let stateIndex = 0, refIndex = 0, mounted = false;
    const states = [], refs = [];
    const React = {
        useRef: current => { const index = refIndex++; return refs[index] ||= { current }; },
        useMemo: fn => fn(),
        useState: initial => { const index = stateIndex++; if (!(index in states)) states[index] = initial; return [states[index], value => states[index] = value]; },
        useEffect: fn => { if (!mounted) { const cleanup = fn(); if (cleanup) cleanups.push(cleanup); } },
    };
    const nodesRef = { current: fixture() }, selectedRef = { current: new Set(["a", "b"]) };
    const commits = [], drags = [];
    let enabled = true;
    let ui;
    const render = () => {
        stateIndex = refIndex = 0;
        ui = useCanvasGroups(React, {
            nodes: nodesRef.current, selected: selectedRef.current, nodesRef, selectedRef,
            setSelected: ids => { selectedRef.current = ids; },
            commit: nodes => { commits.push(nodes); nodesRef.current = nodes; },
            startDrag: (event, id) => {
                drags.push({ event, id, selected: [...selectedRef.current] });
                selectedRef.current = new Set(selectedRef.current);
            },
            canInteract: () => enabled,
        });
        mounted = true;
    };
    render();
    return { get ui() { return ui; }, render, nodesRef, selectedRef, commits, drags, states,
        enable: value => enabled = value,
        key: patch => {
            const event = { code: "KeyG", altKey: true, target: { closest: () => null }, preventDefault() { this.prevented = true; }, ...patch };
            listeners.get("keydown")(event);
            return event;
        },
        close: () => { cleanups.forEach(fn => fn()); assert.equal(listeners.size, 0); globalThis.window = oldWindow; },
    };
}

test("group shortcut reads live selection, ignores editors and active drag, and cleans up", () => {
    const app = hookHarness();
    try {
        app.key({ target: { closest: () => ({}) } });
        assert.equal(app.commits.length, 0);
        app.enable(false); app.key(); assert.equal(app.commits.length, 0);
        app.enable(true);
        app.selectedRef.current = new Set(["b", "c"]);
        assert.equal(app.key().prevented, true);
        assert.equal(app.nodesRef.current[0].canvasGroup, undefined);
        assert.ok(app.nodesRef.current[1].canvasGroup);
        assert.equal(app.nodesRef.current[1].canvasGroup.id, app.nodesRef.current[2].canvasGroup.id);
    } finally { app.close(); }
});

test("group heading selects its members before delegating drag and never triggers Alt duplication", () => {
    const app = hookHarness();
    try {
        const group = collectCanvasGroups(grouped())[0];
        let prevented = 0;
        app.ui.drag({ button: 0, pointerId: 9, clientX: 32, clientY: 45, altKey: true,
            preventDefault: () => prevented++, stopPropagation() {} }, group);
        assert.deepEqual(app.drags[0].selected, ["a", "b"]);
        assert.equal(app.drags[0].event.altKey, false);
        assert.equal(app.drags[0].event.pointerId, 9);
        assert.equal(app.drags[0].id, "a");
        assert.equal(prevented, 1);
        app.ui.drag({ button: 2 }, group);
        assert.equal(app.drags.length, 1);
    } finally { app.close(); }
});

test("dragging a member after selecting its group moves only that member without removing membership", () => {
    const app = hookHarness();
    try {
        app.nodesRef.current = grouped();
        app.render();
        const group = collectCanvasGroups(app.nodesRef.current)[0];
        app.ui.selectGroup(group);
        app.render();
        app.ui.beforeNodeDrag({ button: 0 }, "a");
        assert.deepEqual([...app.selectedRef.current], ["a"]);
        assert.equal(app.nodesRef.current[0].canvasGroup.id, "group-1");
        assert.equal(app.nodesRef.current[1].canvasGroup.id, "group-1");
        app.ui.selectGroup(group);
        app.render();
        app.ui.beforeNodeDrag({ button: 0, altKey: true }, "a");
        assert.deepEqual([...app.selectedRef.current], ["a", "b"]);
        app.selectedRef.current = new Set(["a", "b", "c"]);
        app.ui.beforeNodeDrag({ button: 0 }, "a");
        assert.deepEqual([...app.selectedRef.current], ["a", "b", "c"], "ordinary multi-selection is unchanged");
    } finally { app.close(); }
});

test("a just-created group also lets the next direct member drag select only that member", () => {
    const app = hookHarness();
    try {
        app.key();
        app.render();
        app.ui.beforeNodeDrag({ button: 0 }, "b");
        assert.deepEqual([...app.selectedRef.current], ["b"]);
    } finally { app.close(); }
});

test("Delete on a selected group removes only its frame and clears selection to protect contents", () => {
    const app = hookHarness();
    try {
        app.key();
        app.render();
        const before = app.nodesRef.current;
        let stopped = false;
        const event = app.key({ code: "Delete", key: "Delete", altKey: false,
            stopImmediatePropagation: () => { stopped = true; } });
        assert.equal(event.prevented, true);
        assert.equal(stopped, true);
        assert.equal(collectCanvasGroups(app.nodesRef.current).length, 0);
        assert.equal(app.nodesRef.current.length, before.length);
        app.nodesRef.current.forEach((node, index) => {
            assert.equal(node.position, before[index].position);
            assert.equal(node.metadata, before[index].metadata);
        });
        assert.equal(app.selectedRef.current.size, 0);
    } finally { app.close(); }
});

test("Delete of individual or ordinary multi-selected nodes is left to existing canvas behavior", () => {
    const app = hookHarness();
    try {
        app.key();
        app.render();
        app.ui.beforeNodeDrag({ button: 0 }, "a");
        let stopped = false;
        const event = app.key({ code: "Delete", key: "Delete", altKey: false,
            stopImmediatePropagation: () => { stopped = true; } });
        assert.equal(event.prevented, undefined);
        assert.equal(stopped, false);
        assert.equal(collectCanvasGroups(app.nodesRef.current).length, 1);
        app.selectedRef.current = new Set(["a", "b"]);
        assert.equal(app.key({ code: "Backspace", key: "Backspace", altKey: false }).prevented, undefined);
    } finally { app.close(); }
});

test("member dragging stays independent after the shared drag engine replaces the selection Set", () => {
    const app = hookHarness();
    try {
        app.nodesRef.current = grouped();
        app.render();
        app.ui.drag({ button: 0, pointerId: 7, clientX: 10, clientY: 20,
            preventDefault() {}, stopPropagation() {} }, collectCanvasGroups(app.nodesRef.current)[0]);
        app.render();
        app.ui.beforeNodeDrag({ button: 0 }, "b");
        assert.deepEqual([...app.selectedRef.current], ["b"]);
    } finally { app.close(); }
});

test("group background supports dragging, Shift preserves marquee, and node layer stays above groups", () => {
    const React = { createElement: (type, props, ...children) => ({ type, props: { ...props, children } }) };
    const { Layer } = createCanvasGroupComponents(React, {});
    const calls = [];
    const ui = { groups: collectCanvasGroups(grouped()), selected: new Set(),
        drag: event => calls.push(event), selectGroup() {}, openMenu() {}, setEditing() {} };
    const layer = Layer({ ui, visibleIds: new Set(["a"]) });
    const frame = layer.props.children[0][0];
    const down = { button: 0 };
    frame.props.onPointerDown(down);
    assert.deepEqual(calls, [down]);
    frame.props.onPointerDown({ button: 0, shiftKey: true });
    assert.equal(calls.length, 1);
    assert.match(css, /\.canvas-group-frame\s*\{[^}]*pointer-events:\s*auto/);
    assert.match(css, /\.canvas-group-layer\s*\{[^}]*z-index:\s*-1/);
    assert.ok(bundle.includes("canvasGroupsUI.beforeNodeDrag(event,id);Dh(event,id)"));
});

test("node context-menu actions retain their target even if canvas deselection follows right-button release", () => {
    const app = hookHarness();
    try {
        app.nodesRef.current = grouped();
        app.render();
        app.ui.openNodeMenu({ clientX: 40, clientY: 60, preventDefault() {}, stopPropagation() {} }, "c");
        app.render();
        app.selectedRef.current = new Set(["a"]);
        app.ui.join(collectCanvasGroups(app.nodesRef.current)[0]);
        assert.equal(app.nodesRef.current[2].canvasGroup.id, "group-1");
        app.render();
        app.ui.openNodeMenu({ clientX: 40, clientY: 60, preventDefault() {}, stopPropagation() {} }, "c");
        app.render();
        app.selectedRef.current = new Set(["b"]);
        app.ui.remove();
        assert.equal(app.nodesRef.current[2].canvasGroup, undefined);
        assert.equal(app.nodesRef.current[1].canvasGroup.id, "group-1");
    } finally { app.close(); }
});

test("actual group commit records immediate undo without swallowing a preceding pending edit", () => {
    const start = bundle.indexOf("const commitCanvasGroupChange=");
    const end = bundle.indexOf("const canvasGroupsUI=", start);
    const source = bundle.slice(start, end).replace("const commitCanvasGroupChange=", "");
    const initial = { nodes: fixture(), connections: [], chatSessions: [], activeChatId: null, backgroundMode: "blank", showImageInfo: false };
    const edited = { ...initial, nodes: initial.nodes.map(item => ({ ...item, width: 420 })) };
    const next = createCanvasGroup(edited.nodes, ["a", "b"], () => "g");
    const scope = {
        c: { useCallback: fn => fn }, Ah: () => edited, g: { current: 1 },
        h: { current: initial }, m: { current: { past: [], future: [initial] } },
        vn: { current: edited.nodes }, CANVAS_HISTORY_LIMIT: 25, clearTimeout() {},
        _: value => { scope.value = value; }, mc: value => { scope.controls = value; },
    };
    const commit = vm.runInNewContext(source, scope);
    commit(next);
    assert.equal(scope.m.current.past.length, 2);
    assert.equal(scope.m.current.past[0], initial);
    assert.equal(scope.m.current.past[1], edited);
    assert.equal(scope.h.current.nodes, next);
    assert.equal(scope.vn.current, scope.value);
    assert.equal(scope.m.current.future.length, 0);
    commit(next);
    assert.equal(scope.m.current.past.length, 2);
});

test("bundle integrates both clone paths, group layer ordering, context menu and persistence", () => {
    assert.match(bundle, /Mt=placeCanvasGroupClones\(O\.nodes\.map/);
    assert.match(bundle, /Rn=remapClonedCanvasGroups\(bn\.map/);
    assert.match(bundle, /onContextMenu:K=>canvasGroupsUI\.openNodeMenu\(K,O\.id\)/);
    assert.ok(bundle.indexOf("y.jsx(CanvasGroupLayer,") < bundle.indexOf('className:"canvas-connection-layer'));
    assert.match(css, /canvas-group-layer\s*\{[^}]*z-index:\s*-1/);
    assert.ok(bundle.includes("Cp(vn.current.map(ZM))")); // Existing save pipeline retains complete nodes.
    assert.match(bundle, /_\(!?O\.nodes\)/);
});

test("marquee ignores stale zero-button moves and synchronizes final selected IDs before grouping", () => {
    const start = bundle.indexOf("Bs=c.useCallback(") + 3;
    const end = bundle.indexOf(",Bu=c.useCallback(", start);
    const scope = {
        c: { useCallback: fn => fn },
        ci: { current: { startWorldX: -110, startWorldY: 0, additive: false } },
        vn: { current: fixture() }, oo: { current: new Set() }, Z0: () => false,
        or: (x, y) => ({ x, y }), it: value => { scope.marquee = value; },
        ye: value => { scope.selected = value; },
    };
    const select = vm.runInNewContext(bundle.slice(start, end), scope);
    const initial = scope.ci.current;
    select({ buttons: 0, clientX: 460, clientY: 300 });
    assert.equal(scope.ci.current, initial);
    assert.equal(scope.selected, undefined);
    select({ buttons: 1, clientX: 460, clientY: 300 });
    assert.deepEqual([...scope.oo.current], ["a", "b"]);
    assert.equal(scope.oo.current, scope.selected);
    assert.ok(bundle.includes("if(Q.button===0&&ci.current)Bs({clientX:Q.clientX,clientY:Q.clientY,buttons:1})"));
});

test("actual project load normalization preserves grouping and saved model parameters", () => {
    const start = bundle.indexOf("function n4e(");
    const end = bundle.indexOf("function ig(", start);
    const normalize = vm.runInNewContext(`${bundle.slice(start, end)}; n4e`, {
        Ne: { Image: "image", Text: "text" }, bg: "reverse",
    });
    const nodes = grouped();
    const restored = normalize(plain(nodes));
    assert.deepEqual(plain(restored), plain(nodes));
});

test("actual project.json writer persists names, colors and membership together with nodes and edges", async () => {
    const start = bundle.indexOf("zs=c.useCallback(") + 3;
    const end = bundle.indexOf(",Th=c.useCallback(", start);
    let written;
    const frame = { x: -150, y: -80, width: 900, height: 500 };
    const nodes = updateCanvasGroup(grouped(), "group-1", { name: "References", color: "#4285e8", frame });
    const scope = {
        c: { useCallback: fn => fn }, Cp: async value => value, ZM: value => value,
        Zo: { current: { nextGeneratedImageNumber: 9, root: { getFileHandle: async name => {
            assert.equal(name, "project.json");
            return { createWritable: async () => ({ write: async value => { written = JSON.parse(value); }, close: async () => {} }) };
        } } } },
        vn: { current: nodes }, Jr: { current: [{ id: "edge", fromNodeId: "a", toNodeId: "b" }] },
        Ao: { current: { x: 100, y: 80, k: 0.8 } },
        Y: { title: "Group project" }, Te: "blank", It: false, ne: [], te: null,
        e: { success() {}, error(message) { throw new Error(message); } },
    };
    const save = vm.runInNewContext(bundle.slice(start, end), scope);
    assert.equal(await save(false), true);
    assert.deepEqual(written.nodes, plain(nodes));
    assert.deepEqual(written.connections, scope.Jr.current);
    assert.equal(written.nodes[0].canvasGroup.name, "References");
    assert.equal(written.nodes[1].canvasGroup.color, "#4285e8");
    assert.equal(written.nodes[0].metadata.generationPanelWidth, 1250);
    assert.deepEqual(written.nodes[0].canvasGroup.frame, frame);
});

test("actual canvas paste remaps the copied group, preserves source group, and retains edges", () => {
    const nodes = grouped();
    const start = bundle.indexOf("Cf=c.useCallback(") + 3;
    const end = bundle.indexOf(";c.useCallback(", start);
    const sanitizeStart = bundle.indexOf("function sanitizeCanvasNodeClone(");
    const sanitizeEnd = bundle.indexOf("function zc(", sanitizeStart);
    const data = { nodes: [...nodes], edges: [{ id: "edge", fromNodeId: "a", toNodeId: "b" }] };
    const scope = {
        c: { useCallback: fn => fn }, Ne: { Image: "image" }, ua: "loading", rd: "idle",
        remapClonedCanvasGroups, placeCanvasGroupClones,
        f: { current: { nodes: nodes.slice(0, 2), connections: data.edges } },
        lastPointerWorldRef: { current: { x: 1000, y: 400 } }, Ji: () => ({ x: 0, y: 0 }),
        vn: { current: nodes }, _: fn => { data.nodes = fn(data.nodes); },
        ee: fn => { data.edges = fn(data.edges); },
        ye: ids => { data.selected = ids; }, ve() {}, ut() {}, ln() {},
    };
    const paste = vm.runInNewContext(`${bundle.slice(sanitizeStart, sanitizeEnd)}; (${bundle.slice(start, end)})`, scope);
    assert.equal(paste(), true);
    assert.equal(data.nodes.length, 5);
    const copies = data.nodes.filter(item => data.selected.has(item.id));
    assert.equal(copies.length, 2);
    assert.notEqual(copies[0].canvasGroup.id, "group-1");
    assert.equal(copies[0].canvasGroup.id, copies[1].canvasGroup.id);
    assert.equal(data.nodes[0], nodes[0]);
    assert.equal(data.edges[1].fromNodeId, copies[0].id);
    assert.equal(data.edges[1].toNodeId, copies[1].id);
    assert.equal(copies[0].metadata.model, nodes[0].metadata.model);
    assert.equal(copies[0].metadata.generationPanelWidth, 1250);
});
