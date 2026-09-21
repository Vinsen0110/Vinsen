import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";
import * as mart from "../apimart-api.js";
import * as rh from "../runninghub-api.js";
import * as grsai from "../grsai-api.js";

const source = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const removed = ["gpt-image-2", "gpt-image-2-all", "gpt-image-2-vip", "gpt-image-2-official"];
const constants = Object.fromEntries(
    ["APOLLO_SITE_MODELS", "APOLLO_IMAGE_MODELS", "TUDOU_SITE_MODELS", "TUDOU_IMAGE_MODELS"]
        .map(name => [name, JSON.parse(source.match(new RegExp(`${name}=(\\[[^\\]]*\\])`))[1])]),
);

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

function runtime() {
    const scope = vm.createContext({
        ...constants, ...mart, ...rh, ...grsai,
        DP: "default", VR: "::", a7: "Apilio", cy: "https://api.apilio.ai",
        TUDOU_SITE_ID: "tudou", TUDOU_SITE_NAME: "Tudou", TUDOU_BASE_URL: "https://api.ai-tudou.net",
        zP: site => site,
    });
    for (const name of ["pr", "$S", "ES", "v5", "yx", "CS", "bxe", "hke", "wxe", "xxe", "bd", "siteImageModelNames"]) {
        vm.runInContext(functionSource(name), scope);
    }
    return scope;
}

test("all five site catalogs remove only GPT2 routes and retain their GPT2.5 entries", () => {
    const scope = runtime();
    const channels = scope.wxe();
    assert.equal(channels.length, 5);
    for (const channel of channels) {
        const images = Array.from(scope.siteImageModelNames(channel.id));
        for (const model of removed) {
            assert.ok(!channel.models.includes(model), `${channel.id}: ${model}`);
            assert.ok(!images.includes(model), `${channel.id} image menu: ${model}`);
        }
        assert.equal(images.includes("gpt-image-2.5"), ["default", "runninghub", "apimart"].includes(channel.id));
        assert.ok(images.includes("nano-banana-pro"));
    }
});

test("saved settings are migrated to new catalogs without removing credentials", () => {
    const scope = runtime();
    const old = Array.from(scope.wxe(), site => ({
        ...site, apiKey: `${site.id}-test`, models: [...site.models, ...removed],
    }));
    const channels = scope.wxe("", old);
    for (const channel of channels) {
        assert.equal(channel.apiKey, `${channel.id}-test`);
        assert.equal(scope.yx(`${channel.id}::gpt-image-2`, channels, channel.id), "");
        assert.equal(scope.xxe({ channels }, `${channel.id}::gpt-image-2`), false);
    }
});

test("old project models cannot reappear as selectable options or masquerade as another model", () => {
    const scope = runtime();
    const kept = ["default::nano-banana-pro", "default::gpt-image-2.5"];
    const options = [...kept, ...removed.map(model => `default::${model}`)];
    const config = { imageModels: options, models: options };
    assert.deepEqual(Array.from(scope.CS(config, "image")), kept);
    assert.deepEqual(Array.from(scope.CS(config)), kept);
    for (const model of removed) {
        assert.equal(scope.hke(`default::${model}`, kept), `default::${model}`);
    }
    assert.equal(config.imageModels.length, options.length, "do not rewrite historical input");
});

test("GPT2.5 fixed and official variants still build their original Mart and RH requests", () => {
    for (const variant of ["flare", "sunburst"]) {
        const fixed = mart.apiMartImageRequestSpec({
            model: "gpt-image-2.5", gptImageQuality: "fixed", apimartGpt25Variant: variant,
        }, "test");
        assert.equal(fixed.body.model, "gpt-image-2.5-ext");
        assert.equal(fixed.body.version, variant);
        const official = mart.apiMartImageRequestSpec({
            model: "gpt-image-2.5", gptImageQuality: "max", apimartGpt25Variant: variant,
        }, "test");
        assert.equal(official.body.model, `gpt-image-2.5-${variant}`);
        assert.equal(official.body.quality, "max");
        for (const mode of ["fixed", "official"]) {
            const request = rh.runningHubImageRequestSpec({
                model: "gpt-image-2.5", runningHubGpt25Mode: mode, runningHubGpt25Variant: variant,
            }, "test");
            assert.equal(request.endpoint,
                `/openapi/v2/rhart-image-g-2.5${mode === "official" ? "-official-token" : ""}/${variant}/text-to-image`);
        }
    }
});
