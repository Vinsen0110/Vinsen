import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

function between(startMarker, endMarker) {
    const start = bundle.indexOf(startMarker);
    const end = bundle.indexOf(endMarker, start);
    assert.ok(start >= 0 && end > start, `bundle must contain ${startMarker}`);
    return bundle.slice(start, end);
}

const duplicateSource = between("Dh=c.useCallback(", ",Nl=c.useCallback(").slice(3);
const sanitizeSource = between("function sanitizeCanvasNodeClone(", "function zc(");
const plain = value => JSON.parse(JSON.stringify(value));

function image(id, x) {
    return {
        id, type: "image", title: `Original ${id}`,
        position: { x, y: 80 }, width: 360, height: 480,
        metadata: {
            content: `blob:original-${id}`, storageKey: `image:${id}`,
            naturalWidth: 3000, naturalHeight: 4000,
            prompt: "Keep the original full-resolution image",
            model: "test-image-model", generationType: "image", status: "success",
            bytes: 123456, mimeType: "image/png", generatedAt: "2026-09-12T00:00:00Z",
        },
    };
}

function fixture() {
    return {
        nodes: [
            { id: "source", type: "text", position: { x: 0, y: 0 } },
            image("first", 100), image("second", 500),
            { id: "sink", type: "text", position: { x: 900, y: 0 } },
        ],
        connections: [
            { id: "incoming", fromNodeId: "source", toNodeId: "first", label: "reference" },
            { id: "internal", fromNodeId: "first", toNodeId: "second", label: "chain" },
            { id: "outgoing", fromNodeId: "second", toNodeId: "sink", label: "result" },
            { id: "unrelated", fromNodeId: "source", toNodeId: "sink", label: "untouched" },
        ],
    };
}

function harness(selected, data = fixture(), source = duplicateSource) {
    const state = { ...data, selected: new Set(selected), dragging: false, panel: null };
    const queue = [];
    const commits = [];
    let lane = "sync";
    let transitions = 0;
    const setter = key => value => queue.push({
        lane, key, apply: () => {
            state[key] = typeof value === "function" ? value(state[key]) : value;
        },
    });
    const ref = current => ({ current });
    const scope = {
        c: {
            useCallback: fn => fn,
            startTransition: fn => {
                transitions++;
                const previous = lane;
                lane = "transition";
                try { fn(); } finally { lane = previous; }
            },
        },
        Ne: { Image: "image" }, ua: "loading", rd: "idle",
        vn: ref(state.nodes), Jr: ref(state.connections), oo: ref(state.selected),
        T: ref(null), x: ref(false), E: ref(false), Ph: ref([]),
        _: setter("nodes"), ee: setter("connections"), ye: setter("selected"),
        Ii: setter("dragging"), ln: setter("panel"),
        ut: setter("connectionMenu"), He: setter("hoverHandle"),
        un: setter("hoverNode"), ve: setter("selectedConnection"),
    };
    const context = vm.createContext(scope);
    const duplicate = new vm.Script(`${sanitizeSource};(${source})`).runInContext(context);

    function commit() {
        const pending = queue.filter(update => update.lane === "sync");
        for (const update of pending) {
            queue.splice(queue.indexOf(update), 1);
            update.apply();
        }
        // Mirror the canvas layout effect after the urgent event batch commits.
        scope.vn.current = state.nodes;
        scope.Jr.current = state.connections;
        scope.oo.current = state.selected;
        commits.push(plain({ ...state, selected: [...state.selected] }));
    }

    function click(options = {}, nodeId = "first") {
        const calls = { prevented: 0, stopped: 0 };
        duplicate({
            button: 0, altKey: true, shiftKey: false, metaKey: false, ctrlKey: false,
            clientX: 140, clientY: 160, ...options,
            preventDefault: () => calls.prevented++,
            stopPropagation: () => calls.stopped++,
        }, nodeId);
        commit();
        return calls;
    }

    return { state, queue, commits, scope, click, get transitions() { return transitions; } };
}

for (const selected of [["first"], ["first", "second"]]) {
    test(`Alt duplication atomically commits ${selected.length} selected nodes and their connections`, () => {
        const data = fixture();
        const original = plain(data);
        const app = harness(selected, data);
        const calls = app.click();
        const clones = app.state.nodes.slice(data.nodes.length);
        const ids = new Map(clones.map(node => [node.title, node.id]));
        const remap = id => ids.get(`Original ${id}`) || id;

        assert.deepEqual(calls, { prevented: 1, stopped: 1 });
        assert.equal(app.transitions, 0);
        assert.equal(app.queue.length, 0, "no node or connection insertion may await a transition");
        assert.equal(clones.length, selected.length);
        assert.equal(new Set(clones.map(node => node.id)).size, selected.length);
        assert.deepEqual([...app.state.selected], Array.from(clones, node => node.id));
        assert.equal(app.state.dragging, true);
        assert.ok(app.state.selected.has(app.state.panel));
        for (const commit of app.commits) {
            assert.ok(commit.selected.every(id => commit.nodes.some(node => node.id === id)));
            assert.ok(commit.connections.every(edge =>
                commit.nodes.some(node => node.id === edge.fromNodeId) &&
                commit.nodes.some(node => node.id === edge.toNodeId)));
        }
        for (const clone of clones) {
            const source = data.nodes.find(node => node.title === clone.title);
            assert.notEqual(clone.id, source.id);
            assert.notEqual(clone.position, source.position);
            assert.notEqual(clone.metadata, source.metadata);
            assert.deepEqual(plain({ ...clone, id: source.id }), plain(source));
        }
        const copiedEdges = app.state.connections.slice(data.connections.length);
        const incidentEdges = data.connections.filter(edge =>
            selected.includes(edge.fromNodeId) || selected.includes(edge.toNodeId));
        assert.equal(copiedEdges.length, incidentEdges.length);
        copiedEdges.forEach((edge, index) => {
            const source = incidentEdges[index];
            assert.notEqual(edge.id, source.id);
            assert.deepEqual(plain(edge), {
                ...source, id: edge.id,
                fromNodeId: remap(source.fromNodeId), toNodeId: remap(source.toNodeId),
            });
        });
        assert.deepEqual(plain(data), original, "original nodes and stored image data must not change");
    });
}

test("Alt duplication remaps batch children without selecting hidden children", () => {
    const data = fixture();
    const root = data.nodes[1], child = data.nodes[2];
    root.metadata.batchChildIds = [child.id];
    root.metadata.primaryImageId = child.id;
    child.metadata.batchRootId = root.id;
    const app = harness(["first"], data);
    app.click();
    const [clonedRoot, clonedChild] = app.state.nodes.slice(data.nodes.length);
    assert.deepEqual([...app.state.selected], [clonedRoot.id]);
    assert.deepEqual(plain(clonedRoot.metadata.batchChildIds), [clonedChild.id]);
    assert.equal(clonedRoot.metadata.primaryImageId, clonedChild.id);
    assert.equal(clonedChild.metadata.batchRootId, clonedRoot.id);
    assert.deepEqual(
        Array.from(app.scope.T.current.initialSelectedNodes, node => node.id),
        [clonedRoot.id, clonedChild.id],
    );
    assert.equal(app.queue.length, 0);
});

test("a normal left click selects and drags without duplicating or preventing default", () => {
    const app = harness([]);
    const nodes = app.state.nodes, connections = app.state.connections;
    assert.deepEqual(app.click({ altKey: false }), { prevented: 0, stopped: 1 });
    assert.equal(app.state.nodes, nodes);
    assert.equal(app.state.connections, connections);
    assert.deepEqual([...app.state.selected], ["first"]);
    assert.equal(app.state.dragging, true);
    assert.equal(app.transitions, 0);
});

test("Alt with a non-left button does not start duplication or dragging", () => {
    const app = harness([]);
    const nodes = app.state.nodes;
    assert.deepEqual(app.click({ button: 2 }), { prevented: 0, stopped: 0 });
    assert.equal(app.state.nodes, nodes);
    assert.equal(app.state.selected.size, 0);
    assert.equal(app.state.dragging, false);
});

test("the scheduler exposes the old deferred-insertion selection mismatch", () => {
    const insertion = "_(Vt=>[...Vt,...Rn]),nt.length&&ee(Vt=>[...Vt,...nt])";
    assert.equal(duplicateSource.split(insertion).length - 1, 1);
    const deferred = duplicateSource.replace(insertion, `c.startTransition(()=>{${insertion}})`);
    const app = harness(["first"], fixture(), deferred);
    app.click();
    assert.equal(app.transitions, 1);
    assert.ok(app.queue.some(update => update.key === "nodes"));
    assert.ok(app.queue.some(update => update.key === "connections"));
    assert.ok([...app.state.selected].some(id => !app.state.nodes.some(node => node.id === id)));
});
