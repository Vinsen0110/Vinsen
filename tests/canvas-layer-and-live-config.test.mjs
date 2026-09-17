import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const html = await readFile(new URL("../index.html", import.meta.url), "utf8");

function block(start, end) {
    const from = bundle.indexOf(start);
    const to = bundle.indexOf(end, from + start.length);
    assert.ok(from >= 0 && to > from, start);
    return bundle.slice(from, to);
}

test("connections and all nodes occupy separate stacking contexts", () => {
    assert.match(bundle, /className:"canvas-connection-layer[^"]*",style:\{pointerEvents:"none",zIndex:0\}/);
    assert.match(bundle, /className:"canvas-node-layer",children:m1\.map/);
    assert.match(html, /\.canvas-node-layer\s*\{[^}]*z-index: 1;[^}]*isolation: isolate;/s);
    assert.match(bundle, /contain:"layout style",background:D\.canvas\.background,borderRadius:24/);
    assert.doesNotMatch(bundle, /pointerEvents:"none",transform:"translateZ\(0\)",zIndex:0/);
});

test("memoized nodes invalidate cached request callbacks when settings change", () => {
    const equal = vm.runInNewContext(`${block("function Fze(", "function Hze(")};Fze`, {
        Hze: () => true, cX: () => true, Vze: () => true,
    });
    const before = { data: {}, requestConfig: { apiKey: "" } };
    assert.equal(equal(before, { ...before }), true);
    assert.equal(equal(before, { ...before, requestConfig: { apiKey: "new-key" } }), false);
    assert.match(bundle, /data:O,requestConfig:N,/);
});

for (const retry of [false, true]) {
    test(`${retry ? "retry" : "generate"} reads credentials synchronously from the current store`, async () => {
        const old = { apiKey: "", imageModel: "tudou::nano-banana-pro", channels: [{ id: "tudou", apiKey: "" }] };
        const latest = { ...old, apiKey: "latest-key", channels: [{ id: "tudou", apiKey: "latest-key" }] };
        const warnings = [];
        let checked;
        const scope = {
            c: { useCallback: fn => fn },
            Sa: { current: new Set() }, vn: { current: [] }, Jr: { current: [] },
            N: old, Xr: { getState: () => ({ config: latest }) },
            Q0: config => ({ ...config, model: config.imageModel }),
            r4e: () => null, Ne: { Image: "image", Text: "text", Video: "video", Audio: "audio" },
            e: { warning: message => warnings.push(message) },
            R(config) { checked = config; return Boolean(config.channels[0].apiKey); },
        };
        const source = retry
            ? block("const Tf=c.useCallback(async O=>", "const Ht=Ca(O.id")
            : block("di=c.useCallback(async(O,K,Q)=>", "dt(O);const Mt=Ca");
        const callback = vm.runInNewContext(`${source.replace(/^const Tf=|^di=/, "")}return Se;})`, scope);
        const result = retry ? await callback({ id: "n", type: "image" }) : await callback("n", "image", "test");
        assert.equal(checked.channels[0].apiKey, "latest-key");
        assert.equal(result.apiKey, "latest-key");
        assert.equal(result.model, "tudou::nano-banana-pro");
        assert.equal(warnings.length, 0);
        assert.equal(old.apiKey, "");
    });
}

test("a delayed billing lookup cannot restore an old key or active-key selection", async () => {
    let finish;
    let writes = [];
    let site = { id: "default", activeKeyId: "key-1", apiKey: "old-key", apiKeys: [{ id: "key-1", value: "old-key" }] };
    const scope = {
        g: { provider: "apilio", baseUrl: "https://example.test" },
        activeKey: { id: "key-1", value: "old-key" }, v: "default",
        billingRequestRef: { current: 0 },
        setBillingDetecting() {}, setBillingError() {},
        detectApolloKeyBillingGroup: () => new Promise(resolve => { finish = resolve; }),
        Xr: { getState: () => ({ config: { channels: [site] } }) },
        normalizeSiteApiKeys: value => value,
        updateSite: patch => writes.push(patch),
    };
    const source = block("detectActiveKeyBilling=async", "\nC=()=>").trim().replace(/,$/, "");
    const detect = vm.runInNewContext(`(${source.replace("detectActiveKeyBilling=", "")})`, scope);
    const pending = detect();
    site = { ...site, apiKey: "new-key", apiKeys: [{ id: "key-1", value: "new-key" }] };
    finish({ group: "default", source: "test", checkedAt: "now" });
    await pending;
    assert.equal(writes.length, 0);
    const switched = detect();
    site = { ...site, activeKeyId: "key-2", apiKey: "other-key" };
    finish({ group: "default" });
    await switched;
    assert.equal(writes.length, 0);
});
