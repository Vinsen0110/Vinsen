import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";
import { RUNNINGHUB_TEXT_MODELS } from "../runninghub-api.js";
import { GRSAI_TEXT_MODELS } from "../grsai-api.js";
import { APIMART_SITE_MODELS, APIMART_IMAGE_MODELS, APIMART_TEXT_MODELS } from "../apimart-api.js";

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

const constants = Object.fromEntries(
    ["APOLLO_TEXT_MODELS", "TUDOU_TEXT_MODELS", "APOLLO_SITE_MODELS", "TUDOU_SITE_MODELS"]
        .map(name => [name, JSON.parse(source.match(new RegExp(`${name}=(\\[[^\\]]*\\])`))[1])]),
);
const channels = [
    { id: "default", provider: "apilio", models: constants.APOLLO_SITE_MODELS },
    { id: "runninghub", provider: "runninghub", models: ["nano-banana-pro", ...RUNNINGHUB_TEXT_MODELS] },
    { id: "tudou", provider: "tudou", models: constants.TUDOU_SITE_MODELS },
    { id: "grsai", provider: "grsai", models: ["nano-banana-pro"] },
    { id: "apimart", provider: "apimart", models: APIMART_SITE_MODELS },
].map(site => ({ ...site, apiKey: `${site.id}-test-key`, baseUrl: `https://${site.id}.example`, apiFormat: "openai" }));

function config(activeSiteId, textSiteId = "default") {
    const channel = channels.find(site => site.id === activeSiteId);
    return {
        ...channel, channels, activeSiteId, textSiteId,
        textModel: `${textSiteId}::${{ default: "gemini-3.8-flash", runninghub: "google/gemini-3.7-flash", apimart: "gemini-3.8-flash" }[textSiteId]}`,
        model: `${activeSiteId}::nano-banana-pro`, imageModel: `${activeSiteId}::nano-banana-pro`,
        textModels: activeSiteId === "runninghub" ? ["runninghub::google/gemini-3.7-flash"] : [],
        imageModels: [`${activeSiteId}::nano-banana-pro`],
        quality: "4k", size: "auto",
    };
}

function runtime() {
    const scope = vm.createContext({
        ...constants, RUNNINGHUB_TEXT_MODELS, GRSAI_TEXT_MODELS, APIMART_TEXT_MODELS,
        DP: "default", VR: "::", RUNNINGHUB_SITE_ID: "runninghub", TUDOU_SITE_ID: "tudou",
        GRSAI_SITE_ID: "grsai", APIMART_SITE_ID: "apimart", an: config("default"),
        normalizeSiteApiKeys: site => ({ apiKey: site.apiKey }),
        siteImageModelNames: () => ["nano-banana-pro"],
    });
    for (const name of [
        "ES", "$S", "pr", "Nxe", "siteModelRefs", "siteTextModelNames", "textSiteChannel",
        "normalizeTextConfig", "textRequestConfig", "displayTextModelName", "buildSiteConfigPatch",
        "v5", "yx", "CS", "bxe", "IDe", "ODe", "imageNodeConfig", "cPe", "vke", "Q0",
        "activeSiteChannel", "bd", "xxe",
    ]) vm.runInContext(functionSource(name), scope);
    return scope;
}

test("Apilio, RH, and Mart offer text models; Tudou retains image models", () => {
    const scope = runtime();
    assert.deepEqual(channels.filter(site => scope.siteTextModelNames(site.id).length).map(site => site.id), ["default", "runninghub", "apimart"]);
    assert.deepEqual(constants.TUDOU_SITE_MODELS, ["nano-banana-pro", "gpt-image-2"]);
});

test("Mart text generation passes the real model and credential preflight on every image site", () => {
    const scope = runtime();
    for (const active of channels.map(site => site.id)) {
        const defaults = config(active, "apimart");
        const before = JSON.stringify(defaults);
        for (const name of ["cPe", "vke", "Q0"]) {
            const panel = scope[name](defaults, { metadata: {} }, "text");
            assert.equal(scope.xxe(panel, panel.model), true, `${name}: ${active}`);
            const route = scope.bd(panel, panel.model);
            assert.equal(route.provider, "apimart");
            assert.equal(route.apiKey, "apimart-test-key");
        }
        const missingKey = {
            ...defaults,
            channels: defaults.channels.map(site => site.id === "apimart" ? { ...site, apiKey: "" } : site),
        };
        assert.equal(scope.xxe(missingKey, defaults.textModel), false);
        assert.equal(scope.xxe(defaults, "apimart::unsupported-model"), false);
        assert.equal(JSON.stringify(defaults), before);
    }
    assert.deepEqual(APIMART_IMAGE_MODELS, ["nano-banana-pro", "gpt-image-2", "gpt-image-2.5"]);
    assert.deepEqual(APIMART_SITE_MODELS, [...APIMART_IMAGE_MODELS, ...APIMART_TEXT_MODELS]);
});

test("text panels, picker options and request config agree with global text defaults on every image site", () => {
    const scope = runtime();
    for (const active of channels.map(site => site.id)) {
        for (const textSite of ["default", "runninghub", "apimart"]) {
            const defaults = config(active, textSite);
            const before = JSON.stringify(defaults);
            for (const metadata of [{}, { model: "tudou::gpt-5.5" }, { model: "runninghub::google/gemini-3.7-flash" }]) {
                const node = { metadata: { ...metadata, localPromptPresetId: "builtin-image-reverse" } };
                const original = JSON.stringify(node);
                for (const name of ["cPe", "vke", "Q0"]) {
                    const result = scope[name](defaults, node, "text");
                    assert.equal(result.model, defaults.textModel, `${name}: ${active} / ${textSite}`);
                    assert.equal(result.activeSiteId, textSite);
                    assert.equal(result.apiKey, `${textSite}-test-key`);
                    assert.deepEqual(Array.from(result.textModels), [defaults.textModel]);
                    assert.equal(result.localPromptPresetId, "builtin-image-reverse");
                    assert.equal(scope.textRequestConfig(result).model, result.model);
                    assert.equal(scope.displayTextModelName(result, result.model), textSite === "runninghub" ? "gemini-3.7-flash" : "gemini-3.8-flash");
                }
                assert.equal(JSON.stringify(node), original);
            }
            assert.equal(JSON.stringify(defaults), before, "panel derivation cannot change image defaults");
        }
    }
});

test("RH hides only its display prefix without changing stored or requested model IDs", () => {
    const scope = runtime();
    const defaults = config("default", "runninghub");
    const request = scope.textRequestConfig(defaults);
    assert.equal(scope.displayTextModelName(defaults, defaults.textModel), "gemini-3.7-flash");
    assert.equal(scope.displayTextModelName(request, "google/gemini-3.7-flash"), "gemini-3.7-flash");
    assert.equal(request.model, "runninghub::google/gemini-3.7-flash");
    assert.equal(scope.pr(request.model), "google/gemini-3.7-flash");
    assert.equal(defaults.textModel, "runninghub::google/gemini-3.7-flash");
    assert.equal(scope.displayTextModelName(defaults, "default::google/gemini-3.7-flash"), "google/gemini-3.7-flash");
    assert.equal(scope.displayTextModelName(defaults, "default::gemini-3.8-flash"), "gemini-3.8-flash");
    assert.equal(scope.displayTextModelName(defaults, "tudou::gpt-5.5"), "gpt-5.5");
});

test("legacy Tudou text defaults migrate to Apilio without changing image settings", () => {
    const scope = runtime();
    const legacy = { ...config("tudou"), textSiteId: "tudou", textModel: "tudou::gpt-5.5" };
    const result = scope.normalizeTextConfig(legacy);
    assert.equal(result.textSiteId, "default");
    assert.equal(result.textModel, "default::gemini-3.8-flash");
    for (const key of ["activeSiteId", "imageModel", "quality", "size"]) assert.equal(result[key], legacy[key]);
});

test("switching image site leaves global text selection and new text nodes unchanged", () => {
    const scope = runtime();
    for (const textSite of ["default", "runninghub", "apimart"]) {
        for (const site of channels) {
            const initial = config("default", textSite);
            const next = { ...initial, ...scope.buildSiteConfigPatch(initial, site.id) };
            assert.equal(next.textModel, initial.textModel);
            assert.equal(scope.vke(next, { metadata: {} }, "text").model, initial.textModel);
            assert.equal(scope.Q0(next, { metadata: {} }, "image").model, `${site.id}::nano-banana-pro`);
        }
    }
});
