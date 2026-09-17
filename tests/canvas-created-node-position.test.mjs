import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

function callback(startToken, endToken, scope) {
    const start = bundle.indexOf(startToken);
    const end = bundle.indexOf(endToken, start);
    assert.ok(start >= 0 && end > start);
    return vm.runInNewContext(bundle.slice(start, end), { c: { useCallback: fn => fn }, ...scope });
}

function harness(sourceType = "image") {
    const source = {
        id: "source", type: sourceType, position: { x: 40, y: 60 }, width: 560, height: 400,
        metadata: { content: "existing-reference", storageKey: "local-reference" },
    };
    const state = { nodes: [source], edges: [], selected: null, panel: null };
    const Ne = { Image: "image", Text: "text", Config: "config", Audio: "audio" };
    const sizes = { image: [560, 400], text: [560, 400], config: [340, 240], audio: [340, 120] };
    const start = bundle.indexOf("function zc(");
    const end = bundle.indexOf("function _ke(", start);
    const scope = {
        Ne, N: { channels: [], imageModel: "apimart::nano-banana-pro", textModel: "apimart::gemini-3.8-flash" },
        rd: "idle", da: "success", bg: "reverse prompt",
        vn: { current: state.nodes }, oo: { current: new Set(["source"]) },
        Vr: { image: { width: 560, height: 400 } },
        IA: type => ({ width: sizes[type][0], height: sizes[type][1], title: type, metadata: {} }),
        ODe: (_config, _capability, model) => model,
        pm: () => 1, cn: () => "edge",
        ag: (from, to, _nodes, handle) => handle === "target"
            ? { fromNodeId: to, toNodeId: from } : { fromNodeId: from, toNodeId: to },
        _: update => { state.nodes = update(state.nodes); },
        ee: update => { state.edges = update(state.edges); },
        ye: selected => { state.selected = selected; },
        ln: id => { state.panel = id; },
        ve: () => {}, _e: () => {}, Qo: () => {},
        e: { warning: message => assert.fail(message) },
    };
    scope.zc = vm.runInNewContext(`${bundle.slice(start, end)};zc`, scope);
    return {
        state, scope,
        create: callback("c1=c.useCallback", ",u1=c.useCallback", scope),
        input: callback("u1=c.useCallback", ",fv=c.useCallback", scope),
        world: callback("or=c.useCallback", ",Ru=c.useCallback", {
            i: { current: { getBoundingClientRect: () => ({ left: 120, top: 80 }) } },
            Ao: { current: { x: 170, y: -90, k: 0.37 } },
        }),
    };
}

test("connected nodes use the exact release point for their top-left at any canvas zoom", () => {
    for (const type of ["text", "image", "config", "audio"]) {
        for (const handleType of ["source", "target"]) {
            const app = harness();
            const position = app.world(860, 470);
            const pending = { position, connection: { nodeId: "source", handleType } };
            const original = JSON.stringify(pending);
            app.create(type, pending);
            const node = app.state.nodes[1];
            assert.equal(node.position.x, position.x);
            assert.equal(node.position.y, position.y);
            assert.notEqual(node.position, position);
            assert.equal(JSON.stringify(pending), original);
            assert.equal(app.state.edges.length, 1);
            assert.equal(app.state.selected.has(node.id), true);
            if (type === "text") {
                assert.equal(node.width, 380);
                assert.equal(node.height, 260);
                assert.equal(node.metadata.reversePrompt, true);
                assert.equal(node.metadata.model, "apimart::gemini-3.8-flash");
            }
        }
    }
});

test("text from a text source and image input use the same release-point convention", () => {
    const app = harness("text");
    const pending = { position: { x: -325, y: 1400 }, connection: { nodeId: "source", handleType: "source" } };
    app.create("text", pending);
    assert.equal(app.state.nodes[1].position.x, -325);
    assert.equal(app.state.nodes[1].position.y, 1400);
    assert.equal(app.state.nodes[1].metadata.reversePrompt, undefined);
    const image = harness();
    image.input(pending);
    const node = image.state.nodes[1];
    assert.equal(node.position.x, -325);
    assert.equal(node.position.y, 1400);
    assert.equal(node.width, 560);
    assert.equal(node.metadata.storageKey, "local-reference");
});

test("non-connection creation retains its existing centered placement", () => {
    const app = harness();
    const node = app.scope.zc("text", { x: 800, y: 600 }, {});
    assert.equal(node.position.x, 520);
    assert.equal(node.position.y, 400);
});
