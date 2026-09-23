import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import {
    createCanvasGroup, updateCanvasGroup, collectCanvasGroups, canvasGroupBounds,
    placeCanvasGroupClones, finishCanvasGroupCloneDrag,
} from "../canvas-groups.js";

const node = (id, x, y = 120) => ({ id, type: "image", title: id,
    position: { x, y }, width: 120, height: 100,
    metadata: { model: "gpt-image-2.5", prompt: "@\u56fe\u72471", generationPanelWidth: 1250, content: "blob:original" },
});
const frame = { x: 0, y: 0, width: 600, height: 450 };
const scene = () => updateCanvasGroup(
    createCanvasGroup([node("a", 40), node("b", 300), node("outside", 900)], ["a", "b"], () => "g"),
    "g", { frame });
const copy = (source, x, y = 200) => ({ ...source, id: source.id + "-copy", position: { x, y } });

test("single copy inside its original group retains id, current name, color and manual frame", () => {
    const nodes = scene(), source = nodes[0], clone = copy(source, 200);
    const current = updateCanvasGroup(nodes, "g", { name: "Renamed", color: "#4285e8" });
    const result = placeCanvasGroupClones([clone], [source], current);
    assert.equal(result[0].canvasGroup.id, "g");
    assert.equal(result[0].canvasGroup.name, "Renamed");
    assert.equal(result[0].canvasGroup.color, "#4285e8");
    assert.deepEqual(result[0].canvasGroup.frame, frame);
    assert.deepEqual(result[0].canvasGroup.frameOrigin, clone.position);
    assert.equal(result[0].metadata, source.metadata);
    assert.equal(result[0].position, clone.position);
    assert.equal(result[0].width, source.width);
    assert.equal(result[0].height, source.height);
});

test("single copy dropped outside remains ungrouped, including mere edge overlap", () => {
    const nodes = scene();
    for (const x of [-121, 590, 900]) {
        const [result] = placeCanvasGroupClones([copy(nodes[0], x)], [nodes[0]], nodes);
        assert.equal(result.canvasGroup, undefined);
    }
});

test("outside sources never join even when their copies are positioned over an existing group", () => {
    const nodes = scene();
    assert.equal(placeCanvasGroupClones([copy(nodes[2], 200)], [nodes[2]], nodes)[0].canvasGroup, undefined);
    const moved = nodes.map(item => item.id === "a" ? { ...item, position: { x: 900, y: 120 } } : item);
    assert.equal(placeCanvasGroupClones([copy(moved[0], 200)], [moved[0]], moved)[0].canvasGroup, undefined);
    assert.equal(placeCanvasGroupClones([copy(nodes[0], 200)], [nodes[0]], moved)[0].canvasGroup, undefined,
        "a stale clipboard snapshot must not override a source that is now outside");
});

test("multi-selection only joins eligible inside copies, never unrelated outside nodes", () => {
    const nodes = scene();
    const result = placeCanvasGroupClones([
        copy(nodes[0], 60), copy(nodes[1], 300), copy(nodes[2], 200),
    ], nodes, nodes);
    assert.deepEqual(result.map(item => item.canvasGroup?.id), ["g", "g", undefined]);
    assert.equal(nodes[2].canvasGroup, undefined);
    assert.deepEqual(nodes[0].canvasGroup.frame, frame);
});

test("copying an entire group outside still creates an independent movable group", () => {
    const nodes = scene();
    const result = placeCanvasGroupClones(nodes.slice(0, 2).map(item => copy(item, item.position.x + 800)),
        nodes.slice(0, 2), nodes, () => "new-group");
    assert.deepEqual(result.map(item => item.canvasGroup.id), ["new-group", "new-group"]);
    assert.equal(canvasGroupBounds(result).x, frame.x + 800);
    assert.equal(canvasGroupBounds(result).width, frame.width);
    assert.equal(nodes[0].canvasGroup.id, "g");
});

test("deleted groups cannot be recreated by a single stale clipboard member", () => {
    const nodes = scene();
    const current = nodes.map(({ canvasGroup, ...item }) => item);
    assert.equal(placeCanvasGroupClones([copy(nodes[0], 200)], [nodes[0]], current)[0].canvasGroup, undefined);
});

test("drop reconciliation touches only copied nodes and uses frozen pre-copy auto bounds", () => {
    const nodes = createCanvasGroup([node("a", 0), node("b", 180), node("outside", 900)], ["a", "b"], () => "g");
    const clone = copy(nodes[0], 650);
    const before = [...nodes, clone];
    const result = finishCanvasGroupCloneDrag(before, { scene: nodes, sources: [nodes[0]], copyIds: [clone.id] });
    assert.equal(result[3].canvasGroup, undefined);
    nodes.forEach((item, index) => assert.equal(result[index], item));
    assert.equal(collectCanvasGroups(result)[0].nodes.length, 2);
    assert.equal(finishCanvasGroupCloneDrag(before, null), before);
});

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const pasteStart = bundle.indexOf("Cf=c.useCallback(") + 3;
const pasteEnd = bundle.indexOf(";c.useCallback(", pasteStart);
const sanitizeStart = bundle.indexOf("function sanitizeCanvasNodeClone(");
const sanitizeEnd = bundle.indexOf("function zc(", sanitizeStart);
const plain = value => JSON.parse(JSON.stringify(value));

function pasteHarness(nodes, sources, position) {
    const state = { nodes, connections: [{ id: "edge", fromNodeId: "a", toNodeId: "b" }] };
    const scope = {
        c: { useCallback: fn => fn }, Ne: { Image: "image" }, ua: "loading", rd: "idle",
        placeCanvasGroupClones,
        f: { current: { nodes: sources, connections: state.connections.filter(edge =>
            sources.some(node => node.id === edge.fromNodeId) && sources.some(node => node.id === edge.toNodeId)) } },
        lastPointerWorldRef: { current: position }, Ji: () => ({ x: 0, y: 0 }),
        vn: { current: nodes }, _: fn => { state.nodes = fn(state.nodes); },
        ee: fn => { state.connections = fn(state.connections); },
        ye: ids => { state.selected = ids; }, ve() {}, ut() {}, ln() {},
    };
    const paste = vm.runInNewContext(`${bundle.slice(sanitizeStart, sanitizeEnd)}; (${bundle.slice(pasteStart, pasteEnd)})`, scope);
    assert.equal(paste(), true);
    return state;
}

test("actual canvas paste into a group preserves source membership and model parameters", () => {
    const nodes = scene();
    const state = pasteHarness(nodes, [nodes[0]], { x: 300, y: 300 });
    const clone = state.nodes.find(item => state.selected.has(item.id));
    assert.equal(clone.canvasGroup.id, "g");
    assert.equal(clone.metadata.model, "gpt-image-2.5");
    assert.equal(clone.metadata.generationPanelWidth, 1250);
    assert.equal(clone.metadata.followActiveImageSite, true);
    assert.deepEqual(plain(clone.position), { x: 240, y: 250 });
    assert.deepEqual(plain(state.connections), [{ id: "edge", fromNodeId: "a", toNodeId: "b" }]);
    nodes.forEach((item, index) => assert.equal(state.nodes[index], item));
});

test("actual multi-node paste keeps internal connections while joining the original group", () => {
    const nodes = scene();
    const state = pasteHarness(nodes, nodes.slice(0, 2), { x: 300, y: 300 });
    const clones = state.nodes.filter(item => state.selected.has(item.id));
    assert.deepEqual(Array.from(clones, item => item.canvasGroup.id), ["g", "g"]);
    assert.equal(state.connections.length, 2);
    assert.equal(state.connections[1].fromNodeId, clones[0].id);
    assert.equal(state.connections[1].toNodeId, clones[1].id);
    assert.equal(state.nodes[2], nodes[2]);
});

test("actual paste of an outside node over a group never absorbs it", () => {
    const nodes = scene();
    const state = pasteHarness(nodes, [nodes[2]], { x: 300, y: 300 });
    assert.equal(state.nodes.find(item => state.selected.has(item.id)).canvasGroup, undefined);
    assert.equal(collectCanvasGroups(state.nodes)[0].nodes.length, 2);
});
