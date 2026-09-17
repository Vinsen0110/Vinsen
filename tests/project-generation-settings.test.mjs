import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const source = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

function functionSource(name) {
    const start = source.indexOf(`function ${name}(`);
    assert.notEqual(start, -1, name);
    for (let end = source.indexOf("}", start); end !== -1; end = source.indexOf("}", end + 1)) {
        const candidate = source.slice(start, end + 1);
        try {
            new vm.Script(`(${candidate})`);
            return candidate;
        } catch (error) {
            if (!(error instanceof SyntaxError)) throw error;
        }
    }
    throw new Error(`Cannot extract ${name}`);
}

const channels = [
    { id: "default", name: "Apilio", provider: "apilio", baseUrl: "https://api.apilio.ai", apiKey: "apilio-latest", models: ["nano-banana-pro", "gpt-image-2", "gpt-image-2.5"] },
    { id: "tudou", name: "Tudou", provider: "tudou", baseUrl: "https://api.ai-tudou.net", apiKey: "tudou-latest", models: ["nano-banana-pro", "gpt-image-2"] },
    { id: "apimart", name: "Mart", provider: "apimart", baseUrl: "https://api.apimart.ai", apiKey: "mart-latest", models: ["nano-banana-pro", "gpt-image-2", "gpt-image-2.5"] },
    { id: "runninghub", name: "RH", provider: "runninghub", baseUrl: "https://www.runninghub.ai", apiKey: "rh-latest", models: ["nano-banana-pro", "gpt-image-2", "gpt-image-2.5"] },
    { id: "grsai", name: "Grsai", provider: "grsai", baseUrl: "https://grsai.dakka.com.cn", apiKey: "grsai-latest", models: ["nano-banana-pro", "gpt-image-2-vip"] },
];

function config(site = "default") {
    const channel = channels.find(item => item.id === site);
    return {
        ...channel, channels, activeSiteId: site, imageModel: `${site}::nano-banana-pro`,
        model: `${site}::nano-banana-pro`, imageModels: channel.models.map(model => `${site}::${model}`),
        quality: "4k", size: "auto", gptImageQuality: "medium",
        apimartGpt25Variant: "flare", apimartOutputFormat: "jpeg", apimartBackground: "opaque",
        runningHubGpt25Mode: "fixed", runningHubGpt25Variant: "flare",
        runningHubOutputFormat: "jpeg", runningHubBackground: "opaque",
    };
}

function runtime(extra = {}) {
    const scope = vm.createContext({
        VR: "::", DP: "default", an: config(), Blob, Date,
        c: { useCallback: callback => callback },
        Ne: { Image: "image", Text: "text" },
        normalizeSiteApiKeys: site => ({ apiKey: site.apiKey, apiKeys: site.apiKeys || [] }),
        activeSiteChannel: state => state.channels.find(site => site.id === state.activeSiteId),
        siteImageModelNames: id => channels.find(site => site.id === id)?.models || [],
        IDe: state => state.imageModel,
        textRequestConfig: () => ({ textModel: "default::global-text" }),
        isRunningHubSite: site => site?.provider === "runninghub",
        bg: "reverse prompt",
        ...extra,
    });
    for (const name of [
        "ES", "$S", "pr", "v5", "yx", "Nxe", "siteModelRefs", "CS", "bxe", "ODe",
        "imageNodeConfig", "bd", "xxe", "cPe", "vke", "Q0", "X0", "NX",
        "n4e", "ZM", "Qa", "Hke", "vLocalAsset", "CX", "Ey",
    ]) {
        if (name === "imageNodeConfig" && !source.includes(`function ${name}(`)) continue;
        vm.runInContext(functionSource(name), scope);
    }
    return scope;
}

const metadata = {
    generationType: "generation", model: "runninghub::gpt-image-2.5", quality: "2k", size: "16:9",
    gptImageQuality: "xhigh", runningHubGpt25Mode: "official", runningHubGpt25Variant: "sunburst",
    runningHubOutputFormat: "webp", runningHubBackground: "transparent",
    apimartGpt25Variant: "sunburst", apimartOutputFormat: "png", apimartBackground: "transparent",
    localPromptPresetId: "custom-preset", count: 1,
};
const node = { id: "generated", type: "image", width: 640, height: 360, metadata };
const plain = value => JSON.parse(JSON.stringify(value));

test("restored image panels and request configs retain the saved provider rather than the active site", () => {
    const scope = runtime();
    for (const active of channels.map(site => site.id)) {
        const defaults = config(active);
        const before = JSON.stringify(defaults);
        for (const name of ["vke", "cPe", "Q0"]) {
            const restored = scope[name](defaults, plain(node), "image");
            for (const key of Object.keys(metadata).filter(key => !["generationType", "count"].includes(key))) {
                assert.equal(restored[key], metadata[key], `${active} ${name} ${key}`);
            }
            assert.equal(restored.provider, "runninghub");
            assert.equal(restored.activeSiteId, "runninghub");
            assert.equal(restored.apiKey, "rh-latest");
            assert.ok(restored.imageModels.includes(metadata.model));
        }
        assert.equal(JSON.stringify(defaults), before, "restoring one node must not mutate global defaults");
    }
});

test("same-name models keep their own saved sites and unknown models never become the default", () => {
    const scope = runtime();
    for (const site of channels) {
        const model = `${site.id}::nano-banana-pro`;
        const restored = scope.Q0(config(), { ...node, metadata: { model, quality: "1k" } }, "image");
        assert.equal(restored.model, model);
        assert.equal(restored.apiKey, site.apiKey);
    }
    for (const model of ["missing-site::nano-banana-pro", "tudou::removed-image-model"]) {
        const restored = scope.Q0(config(), { ...node, metadata: { model } }, "image");
        assert.equal(restored.model, model);
        assert.equal(scope.xxe(restored, model), false, "do not silently generate on another route");
    }
    assert.equal(scope.ODe(config(), "text", "tudou::old-text"), "default::global-text");
});

test("generation metadata keeps every image parameter but never credentials", () => {
    const scope = runtime();
    const saved = scope.X0("generation", { ...config("runninghub"), ...metadata }, 1, []);
    for (const key of Object.keys(metadata)) assert.equal(saved[key], metadata[key], key);
    assert.equal(Object.hasOwn(saved, "apiKey"), false);
    assert.equal(Object.hasOwn(saved, "channels"), false);
    assert.equal(Object.hasOwn(saved, "baseUrl"), false);
});

test("new nodes keep current defaults while saved nodes use fresh keys from their original site", () => {
    const scope = runtime();
    const defaults = config("tudou");
    assert.equal(scope.imageNodeConfig(defaults, {}), defaults);
    assert.equal(scope.Q0(defaults, { metadata: {} }, "image").model, "tudou::nano-banana-pro");
    const updated = {
        ...defaults,
        channels: channels.map(site => site.id === "runninghub" ? { ...site, apiKey: "rh-rotated" } : site),
    };
    const restored = scope.Q0(updated, node, "image");
    assert.equal(restored.apiKey, "rh-rotated");
    assert.equal(restored.baseUrl, channels.find(site => site.id === "runninghub").baseUrl);
    assert.equal(restored.activeSiteId, "runninghub");
    assert.equal(defaults.activeSiteId, "tudou");
    assert.equal(node.metadata.model, "runninghub::gpt-image-2.5");
});

test("local project save and reopen preserve actual generated model and parameters", async () => {
    const files = new Map();
    const image = new Blob(["original image bytes"], { type: "image/png" });
    const folder = {
        async getFileHandle(path) {
            return {
                async createWritable() {
                    return { async write(value) { files.set(path, value); }, async close() {} };
                },
                async getFile() { return files.get(path); },
            };
        },
    };
    const savedNode = {
        ...node,
        metadata: { ...metadata, generationPanelWidth: 820, generationPanelHeight: 350, content: "blob:old-session", generatedAt: "2026-09-17T08:30:00Z", naturalWidth: 640, naturalHeight: 360 },
    };
    const scope = runtime({
        Zo: { current: { root: folder, generatedAssets: folder, uploadedAssets: folder, nextGeneratedImageNumber: 1 } },
        vn: { current: [savedNode] }, Jr: { current: [] }, Ao: { current: { x: 1, y: 2, k: 1 } },
        Y: { title: "Round trip" }, Te: "blank", It: false, ne: [], te: null,
        e: { success() {}, error(message) { throw new Error(message); } },
        Sp: async () => image,
        mf: async (blob, kind, name) => {
            files.set(name, blob);
            return `${kind === "generated" ? "\u751f\u6210\u56fe" : "\u4e0a\u4f20\u56fe"}/${name}`;
        },
        storeCanvasImageWithoutDecode: async blob => {
            assert.equal(await blob.text(), "original image bytes");
            return { url: "blob:reopened-session", storageKey: "image:new", width: 640, height: 360, bytes: blob.size, mimeType: blob.type };
        },
    });
    function callback(name, end) {
        const start = source.indexOf(`${name}=c.useCallback`);
        const finish = source.indexOf(end, start);
        assert.ok(start >= 0 && finish > start);
        return vm.runInContext(`(${source.slice(start + name.length + 1, finish)})`, scope);
    }
    scope.Cp = callback("Cp", ",zs=c.useCallback");
    const save = callback("zs", ",Th=c.useCallback");
    assert.equal(await save(false), true);
    const json = files.get("project.json");
    assert.doesNotMatch(json, /apiKey|apiKeys|cloudinary|rh-latest|apilio-latest/);
    const project = JSON.parse(json);
    for (const [key, value] of Object.entries(metadata)) assert.equal(project.nodes[0].metadata[key], value);
    const open = callback("Ep", ",sv=c.useCallback");
    const reopened = await open(project);
    assert.equal(project.nodes[0].metadata.generationPanelWidth, 820);
    assert.equal(project.nodes[0].metadata.generationPanelHeight, 350);
    assert.equal(reopened[0].metadata.generationPanelWidth, 820);
    assert.equal(reopened[0].metadata.generationPanelHeight, 350);
    assert.equal(reopened[0].metadata.content, "blob:reopened-session");
    for (const [key, value] of Object.entries(metadata)) assert.equal(reopened[0].metadata[key], value);
    const restored = scope.Q0(config("tudou"), reopened[0], "image");
    assert.equal(restored.model, metadata.model);
    assert.equal(restored.runningHubGpt25Mode, "official");
    assert.equal(restored.gptImageQuality, "xhigh");
});

test("image retry delegates restoration to the same complete node config as the panel", () => {
    const start = source.indexOf("const Tf=c.useCallback(async O=>");
    const finish = source.indexOf("const Ht=Ca(O.id", start);
    assert.match(source.slice(start, finish), /Q0\(N,\{\.\.\.O,metadata:le\},"image"\)/);
});
