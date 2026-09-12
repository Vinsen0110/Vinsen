import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { APP_STATE_KEYS, createAppStateStorage, createDeferredStateWriter } from "../app-state-storage.js";

const index = await readFile(new URL("../index.html", import.meta.url), "utf8");
const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const settle = () => new Promise(resolve => setImmediate(resolve));

function deferred() {
    let resolve;
    let reject;
    const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
    return { promise, resolve, reject };
}

function mockStorage(overrides = {}) {
    const values = new Map();
    const calls = [];
    return {
        values,
        calls,
        async setItem(key, value) {
            calls.push(["set", key, value]);
            values.set(key, value);
        },
        async removeItem(key) {
            calls.push(["remove", key]);
            values.delete(key);
        },
        async flush() { calls.push(["flush"]); },
        ...overrides,
    };
}

test("application storage retains exactly the four existing Zustand keys", () => {
    assert.deepEqual(APP_STATE_KEYS, [
        "infinite-canvas:ai_config_store",
        "infinite-canvas:theme_store",
        "infinite-canvas:asset_store",
        "infinite-canvas:canvas_store",
    ]);
});

test("native Storage methods are not globally replaced", () => {
    assert.doesNotMatch(index, /Storage\.prototype(?:\.[A-Za-z_$]+|\[[^\]]+\])\s*=/);
    assert.doesNotMatch(index, /Object\.definePropert(?:y|ies)\(\s*Storage\.prototype/);
});

for (const property of ["indexedDB", "localStorage"]) {
    test(`a denied ${property} getter rejects storage readiness without throwing from the factory`, async () => {
        const descriptor = Object.getOwnPropertyDescriptor(globalThis, property);
        const failure = new DOMException(`${property} access denied`, "SecurityError");
        const reports = [];
        let storage;
        let reads = 0;
        try {
            Object.defineProperty(globalThis, property, {
                configurable: true,
                get() {
                    reads += 1;
                    throw failure;
                },
            });
            assert.doesNotThrow(() => {
                storage = createAppStateStorage({
                    ...(property === "indexedDB" ? { localStorage: {} } : { indexedDB: {} }),
                    onError: error => reports.push(error),
                });
            });
        } finally {
            if (descriptor) Object.defineProperty(globalThis, property, descriptor);
            else delete globalThis[property];
        }
        assert.equal(reads, 1);
        assert.equal(typeof storage.flush, "function");
        await assert.rejects(storage.ready, error => error === failure);
        await assert.rejects(storage.flush(), error => error === failure);
        assert.deepEqual(reports, [failure]);
    });
}

test("persisted application stores skip automatic hydration until storage is ready", () => {
    for (const key of APP_STATE_KEYS) {
        const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const alias = bundle.match(new RegExp(`([A-Za-z_$][\\w$]*)\\s*=\\s*"${escapedKey}"`))?.[1];
        const escapedAlias = alias?.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const name = new RegExp(`name:\\s*(?:"${escapedKey}"${escapedAlias ? `|${escapedAlias}(?=[,}])` : ""})`);
        const start = bundle.search(name);
        assert.notEqual(start, -1, `Missing store ${key}`);
        assert.match(bundle.slice(start, start + 900), /skipHydration:\s*(?:true|!0)/, key);
    }
    assert.match(bundle, /(?:await\s+\w+\.ready|\.ready\.then\()/);
    assert.match(bundle, /\.persist\.rehydrate\(/);
});

test("deferred writes coalesce the latest value independently for each key", async () => {
    const storage = mockStorage();
    const writer = createDeferredStateWriter(storage, 60_000);
    writer.setItem("canvas", "old");
    writer.setItem("theme", "dark");
    writer.setItem("canvas", "latest");
    assert.deepEqual(storage.calls, []);
    await writer.flush();
    assert.deepEqual(storage.calls, [
        ["set", "canvas", "latest"],
        ["set", "theme", "dark"],
        ["flush"],
    ]);
});

test("serialization runs only for the latest value in each flushed batch", async () => {
    const storage = mockStorage();
    const serialized = [];
    const writer = createDeferredStateWriter(storage, 60_000, value => {
        serialized.push(value);
        return JSON.stringify(value);
    });
    const obsolete = { state: { projects: [{ id: "old" }] } };
    const latest = { state: { projects: [{ id: "latest" }] } };
    writer.setItem("canvas", obsolete);
    writer.setItem("canvas", latest);
    assert.deepEqual(serialized, []);
    await writer.flush();
    assert.deepEqual(serialized, [latest]);
    assert.deepEqual(storage.calls, [["set", "canvas", JSON.stringify(latest)], ["flush"]]);
});

test("the delay automatically persists the newest queued value", async () => {
    const finished = deferred();
    const storage = mockStorage({ async flush() { finished.resolve(); } });
    const writer = createDeferredStateWriter(storage, 0);
    writer.setItem("canvas", "old");
    writer.setItem("canvas", "new");
    await finished.promise;
    assert.equal(storage.values.get("canvas"), "new");
    assert.deepEqual(storage.calls, [["set", "canvas", "new"]]);
});

test("flush waits for both delayed writes and the storage commit", async () => {
    const write = deferred();
    const commit = deferred();
    const storage = mockStorage({
        async setItem() { await write.promise; },
        async flush() { await commit.promise; },
    });
    const writer = createDeferredStateWriter(storage, 60_000);
    writer.setItem("canvas", "new");
    let completed = false;
    const flushing = writer.flush().then(() => { completed = true; });
    await settle();
    assert.equal(completed, false);
    write.resolve();
    await settle();
    assert.equal(completed, false);
    commit.resolve();
    await flushing;
    assert.equal(completed, true);
});

test("new batches cannot overtake an in-flight write", async () => {
    const first = deferred();
    const storage = mockStorage();
    const setItem = storage.setItem.bind(storage);
    storage.setItem = async (key, value) => {
        if (value === "first") await first.promise;
        await setItem(key, value);
    };
    const writer = createDeferredStateWriter(storage, 60_000);
    writer.setItem("canvas", "first");
    const firstFlush = writer.flush();
    await settle();
    writer.setItem("canvas", "second");
    writer.setItem("canvas", "latest");
    const secondFlush = writer.flush();
    await settle();
    assert.deepEqual(storage.calls, []);
    first.resolve();
    await Promise.all([firstFlush, secondFlush]);
    assert.deepEqual(storage.calls, [
        ["set", "canvas", "first"], ["flush"],
        ["set", "canvas", "latest"], ["flush"],
    ]);
});

test("removing a pending key prevents its later resurrection", async () => {
    const storage = mockStorage();
    const writer = createDeferredStateWriter(storage, 60_000);
    writer.setItem("canvas", "obsolete");
    writer.setItem("theme", "dark");
    await writer.removeItem("canvas");
    await writer.flush();
    assert.equal(storage.values.has("canvas"), false);
    assert.equal(storage.values.get("theme"), "dark");
    assert.equal(storage.calls.some(([operation, key]) => operation === "set" && key === "canvas"), false);
});

test("removing a key waits for its in-flight write and leaves it deleted", async () => {
    const pending = deferred();
    const storage = mockStorage();
    const setItem = storage.setItem.bind(storage);
    storage.setItem = async (...args) => {
        await pending.promise;
        await setItem(...args);
    };
    const writer = createDeferredStateWriter(storage, 60_000);
    writer.setItem("canvas", "obsolete");
    const flushing = writer.flush();
    await settle();
    const removing = writer.removeItem("canvas");
    await settle();
    assert.deepEqual(storage.calls, []);
    pending.resolve();
    await Promise.all([flushing, removing]);
    await writer.flush();
    assert.equal(storage.values.has("canvas"), false);
});

test("flush also waits for a removal already in progress", async () => {
    const pending = deferred();
    const storage = mockStorage({ async removeItem() { await pending.promise; } });
    const writer = createDeferredStateWriter(storage, 60_000);
    const removing = writer.removeItem("canvas");
    let completed = false;
    const flushing = writer.flush().then(() => { completed = true; });
    await settle();
    const completedBeforeRemoval = completed;
    pending.resolve();
    await Promise.all([removing, flushing]);
    assert.equal(completedBeforeRemoval, false);
});

test("flush rejects when a delayed write fails", async () => {
    const pending = deferred();
    const failure = new Error("write failed");
    const writer = createDeferredStateWriter(mockStorage({
        async setItem() { await pending.promise; },
    }), 60_000);
    writer.setItem("canvas", "new");
    const rejected = assert.rejects(writer.flush(), error => error === failure);
    pending.reject(failure);
    await rejected;
});

test("flush propagates a storage commit error", async () => {
    const failure = new Error("transaction aborted");
    const writer = createDeferredStateWriter(mockStorage({
        async flush() { throw failure; },
    }), 60_000);
    writer.setItem("canvas", "new");
    await assert.rejects(writer.flush(), error => error === failure);
});

test("an earlier rejected batch does not block a later successful write", async () => {
    const failure = new Error("temporary write error");
    const storage = mockStorage();
    const setItem = storage.setItem.bind(storage);
    storage.setItem = async (key, value) => {
        if (value === "old") throw failure;
        await setItem(key, value);
    };
    const writer = createDeferredStateWriter(storage, 60_000);
    writer.setItem("canvas", "old");
    await assert.rejects(writer.flush(), error => error === failure);
    writer.setItem("canvas", "new");
    await writer.flush();
    assert.equal(storage.values.get("canvas"), "new");
});
