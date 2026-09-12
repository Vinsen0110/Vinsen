import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const assetKey = "infinite-canvas:asset_store";

function assetAdapter(assets, overrides = {}) {
    const writes = [];
    const value = JSON.stringify({ state: { assets }, version: 0 });
    const start = bundle.indexOf("B$e={hydratedAssets:");
    const end = bundle.indexOf(",gs=fh()", start);
    assert.ok(start >= 0 && end > start);
    const context = {
        Lm: {
            getItem: async () => value,
            setItem: async (key, text) => { writes.push({ key, text }); },
            removeItem: async () => {},
        },
        Xl: async key => `blob:current-image-${key}`,
        Og: async key => `blob:current-video-${key}`,
        bo: async () => ({
            storageKey: "image:migrated", url: "blob:migrated",
            bytes: 10, mimeType: "image/png",
        }),
        ...overrides,
    };
    const adapter = vm.runInNewContext(`const ${bundle.slice(start, end)};B$e`, context);
    return { adapter, writes, value };
}

test("asset hydration restores runtime URLs without persisting hydration-only changes", async () => {
    const assets = [
        { id: "image", kind: "image", coverUrl: "blob:old-image", data: { storageKey: "image:one", dataUrl: "blob:old-image" } },
        { id: "video", kind: "video", coverUrl: "", data: { storageKey: "video:one", url: "blob:old-video" } },
        { id: "text", kind: "text", coverUrl: "", data: { content: "Keep text" } },
    ];
    for (let tab = 0; tab < 2; tab++) {
        const { adapter, writes } = assetAdapter(assets);
        const hydrated = await adapter.getItem(assetKey);
        assert.equal(hydrated.state.assets[0].data.dataUrl, "blob:current-image-image:one");
        assert.equal(hydrated.state.assets[1].data.url, "blob:current-video-video:one");
        await adapter.setItem(assetKey, { state: { assets: hydrated.state.assets }, version: 0 });
        assert.equal(writes.length, 0);
        await adapter.setItem(assetKey, {
            state: { assets: hydrated.state.assets.map(asset => ({ ...asset, title: "Edited" })) },
            version: 0,
        });
        assert.equal(writes.length, 1);
        assert.equal(JSON.parse(writes[0].text).state.assets[0].title, "Edited");
    }
});

test("legacy data-image conversion still persists new durable storage keys", async () => {
    const { adapter, writes } = assetAdapter([
        { id: "legacy", kind: "image", coverUrl: "data:image/png;base64,A", data: { dataUrl: "data:image/png;base64,A" } },
    ]);
    const hydrated = await adapter.getItem(assetKey);
    await adapter.setItem(assetKey, hydrated);
    assert.equal(writes.length, 1);
    assert.equal(JSON.parse(writes[0].text).state.assets[0].data.storageKey, "image:migrated");
});

test("failed durable image conversion preserves legacy data without a hydration rewrite", async () => {
    const source = "data:image/png;base64,A";
    const { adapter, writes, value } = assetAdapter([
        { id: "legacy", kind: "image", coverUrl: source, data: { dataUrl: source } },
    ], { bo: async () => ({ storageKey: "", url: source, bytes: 10, mimeType: "image/png" }) });
    const hydrated = await adapter.getItem(assetKey);
    await adapter.setItem(assetKey, hydrated);
    assert.equal(writes.length, 0);
    assert.equal(JSON.parse(value).state.assets[0].data.dataUrl, source);
});

function cleanupHarness(error) {
    const events = [];
    let scheduled;
    const start = bundle.indexOf("cleanupImages:n=>") + "cleanupImages:".length;
    const end = bundle.indexOf("},0)}", start) + "},0)}".length;
    assert.ok(start >= "cleanupImages:".length && end > start);
    const context = {
        window: { setTimeout: callback => { scheduled = callback; } },
        ENe: { flush: async () => { events.push("flush"); if (error) throw error; } },
        uq: async callback => callback(),
        $Ne: { useCanvasStore: { getState: () => ({ projects: [] }) } },
        t: () => ({ assets: [] }),
        cq: async () => { events.push("images"); },
        D$e: async () => { events.push("media"); },
        showStorageError: failure => { events.push(failure); },
    };
    const source = bundle.slice(start, end).replaceAll("import.meta.url", JSON.stringify(import.meta.url));
    const cleanup = vm.runInNewContext(`(${source})`, context);
    return { events, run: async () => { cleanup({ assets: [] }); await scheduled(); } };
}

test("asset cleanup waits for app persistence before deleting image or media blobs", async () => {
    const { events, run } = cleanupHarness();
    await run();
    assert.deepEqual(events, ["flush", "images", "media"]);
});

test("asset cleanup reports persistence failure and preserves every blob", async () => {
    const error = Object.assign(new Error("Newer state belongs to another tab"), { name: "StorageConflictError" });
    const { events, run } = cleanupHarness(error);
    await run();
    assert.deepEqual(events, ["flush", error]);
});

test("canvas node autosave waits for both hydration and an opened project directory", () => {
    assert.ok(bundle.includes(
        "c.useEffect(()=>{!We||!Me||x.current||I(a,{nodes:B,connections:X,chatSessions:ne,activeChatId:te,backgroundMode:Te,showImageInfo:It})},[te,Te,ne,X,B,a,We,Me,It,I])",
    ));
});

test("canvas viewport autosave does not schedule writes before a project directory opens", () => {
    assert.ok(bundle.includes(
        "c.useEffect(()=>{if(We&&Me)return v.current&&clearTimeout(v.current),v.current=setTimeout(()=>{I(a,{viewport:Ao.current}),v.current=null},500),()=>{v.current&&clearTimeout(v.current)}},[a,We,Me,I,oe])",
    ));
});
