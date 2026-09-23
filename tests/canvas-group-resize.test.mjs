import assert from "node:assert/strict";
import test from "node:test";
import {
    canvasGroupBounds, collectCanvasGroups, createCanvasGroup, updateCanvasGroup,
    dissolveCanvasGroup, remapClonedCanvasGroups, moveCanvasGroupFrames,
    resizeCanvasGroupFrame, beginCanvasGroupResize,
} from "../canvas-groups.js";

const fixture = () => createCanvasGroup([
    { id: "a", position: { x: 0, y: 100 }, width: 200, height: 100, metadata: { model: "gpt-image-2.5" } },
    { id: "b", position: { x: 300, y: 100 }, width: 200, height: 100, metadata: { model: "nano-banana-pro" } },
    { id: "other", position: { x: 700, y: 100 }, width: 200, height: 100 },
], ["a", "b"], () => "g");
const frame = { x: -50, y: 20, width: 700, height: 400 };
const manual = () => updateCanvasGroup(fixture(), "g", { frame });

test("manual frames keep exact dimensions through rename, color and JSON roundtrip", () => {
    const nodes = manual();
    const changed = updateCanvasGroup(nodes, "g", { name: "References", color: "#4285e8" });
    const saved = JSON.parse(JSON.stringify(changed));
    assert.deepEqual(canvasGroupBounds(collectCanvasGroups(saved)[0].nodes), frame);
    assert.deepEqual(saved[0].canvasGroup.frame, saved[1].canvasGroup.frame);
    assert.equal(changed[0].position, nodes[0].position);
    assert.equal(changed[0].metadata, nodes[0].metadata);
    assert.equal(updateCanvasGroup(changed, "g", { frame: { ...frame } }), changed);
    assert.equal(changed[2], nodes[2]);
});

test("malformed saved frames fall back to automatic bounds", () => {
    const nodes = fixture();
    for (const broken of [{ ...frame, width: 0 }, { ...frame, y: NaN }, { ...frame, height: Infinity }]) {
        const changed = updateCanvasGroup(nodes, "g", { frame: broken });
        assert.deepEqual(canvasGroupBounds(changed.slice(0, 2)), canvasGroupBounds(nodes.slice(0, 2)));
    }
});

test("eight resize directions preserve opposite edges and minimum frame dimensions", () => {
    const expected = {
        n: { x: -50, y: 60, width: 700, height: 360 },
        ne: { x: -50, y: 60, width: 750, height: 360 },
        e: { x: -50, y: 20, width: 750, height: 400 },
        se: { x: -50, y: 20, width: 750, height: 440 },
        s: { x: -50, y: 20, width: 700, height: 440 },
        sw: { x: 0, y: 20, width: 650, height: 440 },
        w: { x: 0, y: 20, width: 650, height: 400 },
        nw: { x: 0, y: 60, width: 650, height: 360 },
    };
    for (const direction of Object.keys(expected)) {
        assert.deepEqual(resizeCanvasGroupFrame(frame, direction, 50, 40), expected[direction]);
    }
    assert.deepEqual(resizeCanvasGroupFrame(frame, "nw", 9000, 9000),
        { x: 410, y: 320, width: 240, height: 100 });
    assert.deepEqual(resizeCanvasGroupFrame(frame, "se", -9000, -9000),
        { x: -50, y: 20, width: 240, height: 100 });
});

test("moving one member leaves manual frame and other members in place", () => {
    const before = manual();
    const moved = before.map(node => node.id === "a" ? { ...node, position: { x: -900, y: 300 } } : node);
    const after = moveCanvasGroupFrames(before, moved);
    assert.deepEqual(canvasGroupBounds(after.slice(0, 2)), frame);
    assert.deepEqual(after[0].position, { x: -900, y: 300 });
    assert.equal(after[1].position, before[1].position);
    assert.equal(after[2], before[2]);
    assert.equal(after[0].metadata, before[0].metadata);
    assert.deepEqual(after[0].canvasGroup.frameOrigin, after[0].position);
});

test("moving all members translates manual frame once without changing dimensions", () => {
    const before = manual();
    const moved = before.map(node => node.id !== "other" ?
        { ...node, position: { x: node.position.x + 40, y: node.position.y - 60 } } : node);
    const after = moveCanvasGroupFrames(before, moved);
    assert.deepEqual(canvasGroupBounds(after.slice(0, 2)), { ...frame, x: -10, y: -40 });
    assert.deepEqual(moveCanvasGroupFrames(after, after)[0].canvasGroup.frame, after[0].canvasGroup.frame);
});

test("manual frame copies translate to pasted members and get an independent group id", () => {
    const nodes = manual();
    const copies = nodes.slice(0, 2).map(node => ({
        ...node, id: node.id + "-copy", position: { x: node.position.x + 1000, y: node.position.y - 300 },
    }));
    const cloned = remapClonedCanvasGroups(copies, () => "copy");
    assert.deepEqual(canvasGroupBounds(cloned), { ...frame, x: 950, y: -280 });
    assert.equal(cloned[0].canvasGroup.id, "copy");
    assert.deepEqual(cloned[0].canvasGroup.frameOrigin, cloned[0].position);
    assert.deepEqual(nodes[0].canvasGroup.frame, frame);
});

test("deleting a manual group frame retains every node, position and provider setting", () => {
    const nodes = manual();
    const after = dissolveCanvasGroup(nodes, "g");
    assert.equal(after.length, nodes.length);
    assert.equal(collectCanvasGroups(after).length, 0);
    after.forEach((node, index) => {
        assert.equal(node.id, nodes[index].id);
        assert.equal(node.position, nodes[index].position);
        assert.equal(node.metadata, nodes[index].metadata);
        assert.equal(node.width, nodes[index].width);
        assert.equal(node.height, nodes[index].height);
    });
});

function resizeHarness(scale = 1, direction = "se") {
    const listeners = new Map(), frames = new Map(), commits = [];
    let ended = 0, sequence = 0;
    const element = { style: {}, setAttribute() {}, removeAttribute() {} };
    const event = (overrides = {}) => ({ pointerId: 7, button: 0, buttons: 1,
        clientX: 100, clientY: 200, preventDefault() {}, stopPropagation() {},
        currentTarget: { closest: () => element }, ...overrides });
    const host = {
        addEventListener: (type, fn) => listeners.set(type, fn),
        removeEventListener: type => listeners.delete(type),
        requestAnimationFrame: fn => { const id = ++sequence; frames.set(id, fn); return id; },
        cancelAnimationFrame: id => frames.delete(id),
    };
    const cancel = beginCanvasGroupResize(event(), collectCanvasGroups(manual())[0], direction, scale,
        value => commits.push(value), () => ended++, host);
    return {
        commits, element, frames, listeners, cancel, get ended() { return ended; },
        dispatch: (type, overrides = {}) => listeners.get(type)?.(event(overrides)),
        flush: () => { for (const [id, fn] of frames) { frames.delete(id); fn(); } },
    };
}

for (const scale of [0.5, 1, 2]) {
    test(`resize preview batches frames and commits only final release at zoom ${scale}`, () => {
        const app = resizeHarness(scale);
        app.dispatch("pointermove", { clientX: 120, clientY: 230 });
        app.dispatch("pointermove", { clientX: 150, clientY: 240 });
        assert.equal(app.frames.size, 1);
        assert.equal(app.commits.length, 0);
        app.flush();
        assert.equal(app.element.style.width, `${700 + 50 / scale}px`);
        app.dispatch("pointerup", { buttons: 0, clientX: 180, clientY: 260 });
        assert.deepEqual(app.commits, [{ ...frame, width: 700 + 80 / scale, height: 400 + 60 / scale }]);
        assert.equal(app.ended, 1);
        assert.equal(app.listeners.size, 0);
        assert.equal(app.frames.size, 0);
    });
}

test("resize supports release without move and ignores foreign pointers", () => {
    const app = resizeHarness();
    app.dispatch("pointermove", { pointerId: 9, clientX: 900 });
    app.dispatch("pointerup", { pointerId: 9, buttons: 0, clientX: 900 });
    app.dispatch("pointercancel", { pointerId: 9 });
    assert.equal(app.ended, 0);
    assert.equal(app.frames.size, 0);
    app.dispatch("pointerup", { buttons: 0, clientX: 170, clientY: 220 });
    assert.deepEqual(app.commits, [{ ...frame, width: 770, height: 420 }]);
});

for (const type of ["pointercancel", "blur", "keydown", "unmount"]) {
    test(`resize ${type} cancels without history and restores original bounds`, () => {
        const app = resizeHarness();
        app.dispatch("pointermove", { clientX: 180 });
        app.flush();
        if (type === "unmount") app.cancel();
        else app.dispatch(type, { key: "Escape", stopImmediatePropagation() {} });
        assert.equal(app.element.style.width, "700px");
        assert.equal(app.commits.length, 0);
        assert.equal(app.ended, 1);
        assert.equal(app.listeners.size, 0);
    });
}
