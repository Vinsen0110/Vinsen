import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";
import { isRunningHubSite, runningHubImagePrice } from "../runninghub-api.js";
import { isApiMartSite, apiMartImagePrice } from "../apimart-api.js";
import { isGrsaiSite } from "../grsai-api.js";

const source = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

function functionSource(name) {
    const start = source.indexOf(`function ${name}(`);
    assert.notEqual(start, -1, `${name} must exist in the shipped bundle`);
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

const context = vm.createContext({
    VR: "::",
    hX: ["auto", "1k", "2k", "4k"].map(value => ({ value })),
    isRunningHubSite,
    runningHubImagePrice,
    isApiMartSite,
    apiMartImagePrice,
    isGrsaiSite,
    bd: (config, model) => config.channels.find(channel => channel.id === model.split("::")[0]),
});
for (const name of [
    "$S", "pr", "jMe", "isTudouSite", "isApilioSite",
    "tudouResolution", "tudouQuality", "tudouImagePrice", "vX", "pke",
]) {
    vm.runInContext(functionSource(name), context);
}

test("Tudou Nano Banana Pro uses the supplied 2026-09-24 default-group prices", () => {
    for (const model of [
        "nano-banana-pro", "tudou::nano-banana-pro", "nano-banana-pro-2k",
        "nano-banana-pro-4k", "gemini-3-pro-image-preview",
    ]) {
        for (const [quality, price] of [["auto", 0.1], ["1k", 0.1], ["2k", 0.1], ["4k", 0.14]]) {
            assert.equal(context.tudouImagePrice({ model, quality }), price, `${model} ${quality}`);
            assert.equal(context.tudouImagePrice({ imageModel: model, quality }), price);
        }
    }
});

test("Tudou resolution normalization preserves new prices for saved legacy settings", () => {
    for (const quality of ["low", "standard", "medium", "hd", " 2K ", "auto", undefined]) {
        assert.equal(context.tudouImagePrice({ model: "nano-banana-pro", quality }), 0.1);
    }
    for (const quality of ["high", "4K"]) {
        assert.equal(context.tudouImagePrice({ model: "nano-banana-pro", quality }), 0.14);
    }
});

test("canvas estimates use updated unit prices and multiply batch counts correctly", () => {
    for (const [quality, price] of [["auto", 0.1], ["1k", 0.1], ["2k", 0.1], ["4k", 0.14]]) {
        for (const count of [1, 2, 3, 10, 15]) {
            const config = { provider: "tudou", model: "tudou::nano-banana-pro", quality, count };
            assert.equal(context.pke(config), Number((price * count).toFixed(4)));
            assert.equal(context.pke({
                model: config.model, quality, count,
                channels: [{ id: "tudou", provider: "tudou" }],
            }), Number((price * count).toFixed(4)));
        }
    }
});

test("Tudou GPT Image 2 quality-resolution pricing is unchanged", () => {
    const prices = { low: [0.03, 0.04, 0.06], medium: [0.04, 0.06, 0.08], high: [0.06, 0.08, 0.1] };
    for (const model of ["gpt-image-2", "gpt-image-2-all"]) {
        for (const [gptImageQuality, expected] of Object.entries(prices)) {
            ["1k", "2k", "4k"].forEach((quality, index) => {
                assert.equal(context.tudouImagePrice({ model, quality, gptImageQuality }), expected[index]);
            });
        }
    }
    assert.equal(context.tudouImagePrice({ model: "gpt-5.5" }), 0);
});

test("Nano Banana estimates remain isolated from other providers", () => {
    for (const quality of ["1k", "2k", "4k"]) {
        const config = { model: "nano-banana-pro", quality, count: 1 };
        assert.equal(context.pke({ ...config, provider: "grsai" }), 0.18);
        assert.equal(context.pke({ ...config, provider: "apilio" }), quality === "4k" ? 0.55 : 0.4);
        assert.equal(context.pke({ ...config, provider: "runninghub" }), quality === "4k" ? 0.07 : 0.06);
        assert.equal(context.pke({ ...config, provider: "apimart" }), apiMartImagePrice(config));
    }
});
