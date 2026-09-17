import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";
import {
    isRunningHubSite,
    runningHubGpt25Mode,
    runningHubGpt25Variant,
    runningHubImagePrice,
    runningHubImageRequestSpec,
    runningHubQuality,
    runningHubSupportedAspectRatios,
} from "../runninghub-api.js";
import {
    apiMartGptBackground,
    apiMartGptImageQuality,
    apiMartGptOutputFormat,
} from "../apimart-api.js";

const source = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

// Compile the shipped functions instead of recreating their rendering conditions.
function functionSource(name) {
    const start = source.indexOf(`function ${name}(`);
    assert.notEqual(start, -1, `${name} must exist in the shipped canvas bundle`);
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

const plain = value => JSON.parse(JSON.stringify(value));
const jsx = (type, props, key) => ({ type, props: props || {}, key });
const site = provider => config => config?.provider === provider || config?.id === provider;
const context = vm.createContext({
    y: { jsx, jsxs: jsx, Fragment: "Fragment" },
    no: (...values) => values.filter(Boolean).join(" "),
    n$: "Select",
    uke: "ModelSelect",
    LocalImagePromptPresetSelect: "PresetSelect",
    Wwe: "ResolutionIcon",
    Ig: "SettingsIcon",
    VP: "RatioIcon",
    A0: "Label",
    GL: "ChoiceButton",
    Bt: "GenerateButton",
    Ky: "Spinner",
    vM: "PriceIcon",
    kP: "GenerateIcon",
    isRunningHubSite,
    isApiMartSite: site("apimart"),
    isApilioSite: site("apilio"),
    isTudouSite: site("tudou"),
    isGrsaiSite: site("grsai"),
    runningHubGpt25Mode,
    runningHubGpt25Variant,
    runningHubQuality,
    runningHubImagePrice,
    runningHubSupportedAspectRatios,
    apiMartGptBackground,
    apiMartGptImageQuality,
    apiMartGptOutputFormat,
    tudouQuality: config => config.gptImageQuality || "medium",
    activeSiteApiKey: () => null,
    apolloBillingGroupLabel: () => "default",
    hX: [{ value: "auto" }, { value: "1k" }, { value: "2k" }, { value: "4k" }],
    NANO_PRO_RATIO_PRESETS: [{ value: "auto" }],
    orderRatioPresets: values => values,
    jMe: () => false,
    VR: "::",
    v5: value => value,
    DP: "apilio",
    an: {},
    ODe: (_config, _kind, model) => model,
    siteImageModelNames: () => ["nano-banana-pro", "gpt-image-2", "gpt-image-2.5"],
});

const qualityStart = source.indexOf("GPT_IMAGE_QUALITY_OPTIONS=[");
const optionsStart = source.indexOf("const APIMART_GPT_MODE_OPTIONS=");
const ratioStart = source.indexOf("pX=[");
const extraRatioStart = source.indexOf("GPT_IMAGE_EXTRA_RATIO_PRESETS=[");
vm.runInContext(
    `const ${source.slice(qualityStart, source.indexOf(";", qualityStart) + 1)}
${source.slice(optionsStart, source.indexOf(";", optionsStart) + 1)}
const ${source.slice(ratioStart, source.indexOf("],ike=", ratioStart) + 1)};
const ${source.slice(extraRatioStart, source.indexOf(";", extraRatioStart) + 1)}`,
    context,
);
for (const name of [
    "ES", "$S", "pr", "imageNodeConfig", "vX", "gke", "runningHubUiParams", "RunningHub25Controls",
    "runningHub25RatioValue", "runningHub25RatioOptions", "runningHub25ModePatch",
    "apiMartGptMode", "apiMartGpt25Variant", "apiMartOfficialQuality",
    "defaultImageModelParams", "canonicalImageModel", "normalizeImageModelParams",
    "imageModelParamsFromConfig", "projectImageModelParams", "imageGenerationDefaultsKey",
    "imageGenerationDefaultsFor", "updateImageGenerationDefaults", "applyImageGenerationDefaults",
    "cke", "gX", "vke",
]) {
    vm.runInContext(functionSource(name), context);
}
context.bd = (config, model) => {
    const ref = context.$S(model);
    return config.channels.find(channel => channel.id === ref?.channelId)
        || config.channels.find(channel => channel.id === config.activeSiteId);
};

function nodes(tree) {
    if (tree == null || typeof tree !== "object") return [];
    if (Array.isArray(tree)) return tree.flatMap(nodes);
    const children = typeof tree.type === "function"
        ? tree.type(tree.props)
        : tree.props?.children;
    return [tree, ...nodes(children)];
}

function render(config) {
    const changes = [];
    const tree = context.cke({
        node: { id: "image-node" },
        config,
        canSubmit: true,
        credits: runningHubImagePrice(config),
        onConfigChange: (id, patch) => changes.push({ id, patch: plain(patch) }),
        onGenerate() {},
        onStop() {},
    });
    const all = nodes(tree);
    return {
        all,
        changes,
        selectors: all.filter(node => node.type === "Select"),
        button: all.find(node => node.type === "GenerateButton"),
        select(values) {
            return all.find(node => node.type === "Select"
                && JSON.stringify(node.props.items.map(item => item.value)) === JSON.stringify(values));
        },
    };
}

const rh = {
    provider: "runninghub",
    model: "runninghub::gpt-image-2.5",
    quality: "2k",
    size: "16:9",
    runningHubGpt25Mode: "fixed",
    runningHubGpt25Variant: "flare",
    gptImageQuality: "high",
};
const modes = ["fixed", "official"];
const variants = ["flare", "sunburst"];
const qualities = ["auto", "low", "medium", "high", "xhigh", "max"];
const formats = ["png", "jpeg", "webp"];
const backgrounds = ["auto", "opaque", "transparent"];

test("shipped RH canvas renders mode and variant controls in fixed mode, not official-only fields", () => {
    const view = render(rh);
    assert.equal(view.select(modes)?.props.value, "fixed");
    assert.equal(view.select(variants)?.props.value, "flare");
    assert.equal(view.select(qualities), undefined);
    assert.equal(view.select(formats), undefined);
    assert.equal(view.select(backgrounds), undefined);
    assert.match(JSON.stringify(view.button), /0\.03/);
});

test("mode and Sunburst selections update RH keys and survive a real toolbar rerender", () => {
    const before = render(rh);
    before.select(modes).props.onChange("official");
    before.select(variants).props.onChange("sunburst");
    assert.deepEqual(before.changes, [
        { id: "image-node", patch: { runningHubGpt25Mode: "official" } },
        { id: "image-node", patch: { runningHubGpt25Variant: "sunburst" } },
    ]);
    const changed = Object.assign({}, rh, ...before.changes.map(change => change.patch));
    const after = render(changed);
    assert.equal(after.select(modes).props.value, "official");
    assert.equal(after.select(variants).props.value, "sunburst");
    assert.equal(after.select(qualities).props.value, "high");
    assert.ok(after.select(formats));
    assert.ok(after.select(backgrounds));
    assert.doesNotMatch(JSON.stringify(after.button), /0\.03|PriceIcon|待确认/);
    assert.equal(after.button.props.title, undefined);
    assert.equal(runningHubImageRequestSpec(changed, "draw").endpoint,
        "/openapi/v2/rhart-image-g-2.5-official-token/sunburst/text-to-image");
});

test("RH official edits change RH format/background and preserve Mart's independent values", () => {
    const view = render({
        ...rh,
        runningHubGpt25Mode: "official",
        runningHubOutputFormat: "jpeg",
        apimartOutputFormat: "webp",
        apimartBackground: "opaque",
    });
    view.select(backgrounds).props.onChange("transparent");
    assert.deepEqual(view.changes[0].patch, {
        runningHubBackground: "transparent",
        runningHubOutputFormat: "png",
    });
    view.select(qualities).props.onChange("max");
    assert.deepEqual(view.changes[1].patch, { gptImageQuality: "max" });
    const transparent = render({
        ...rh,
        runningHubGpt25Mode: "official",
        runningHubBackground: "transparent",
        runningHubOutputFormat: "png",
    });
    assert.equal(transparent.select(formats).props.items.find(item => item.value === "jpeg").disabled, true);
});

test("RH toolbar resolves the model's channel even when another site is currently active", () => {
    const view = render({
        ...rh,
        provider: "apimart",
        activeSiteId: "apimart",
        channels: [
            { id: "apimart", provider: "apimart" },
            { id: "runninghub", provider: "runninghub" },
        ],
    });
    view.select(modes).props.onChange("official");
    assert.deepEqual(view.changes[0].patch, { runningHubGpt25Mode: "official" });
});

test("RH mode and variant survive defaults, normalization, projection and switching models", () => {
    const config = {
        ...rh,
        imageModel: rh.model,
        activeSiteId: "runninghub",
        runningHubGpt25Mode: "official",
        runningHubGpt25Variant: "sunburst",
        runningHubOutputFormat: "webp",
        runningHubBackground: "transparent",
        gptImageQuality: "max",
    };
    const params = context.imageModelParamsFromConfig(config);
    const projected = context.projectImageModelParams(params);
    for (const key of ["runningHubGpt25Mode", "runningHubGpt25Variant",
        "runningHubOutputFormat", "runningHubBackground", "gptImageQuality"]) {
        assert.equal(projected[key], config[key], key);
    }
    const away = { ...config, ...context.applyImageGenerationDefaults(config, "gpt-image-2", "runninghub") };
    const back = context.applyImageGenerationDefaults(away, "gpt-image-2.5", "runninghub");
    assert.equal(back.runningHubGpt25Mode, "official");
    assert.equal(back.runningHubGpt25Variant, "sunburst");
    assert.equal(back.gptImageQuality, "max");
    const patch = context.updateImageGenerationDefaults(config, { runningHubGpt25Variant: "flare" });
    assert.equal(patch.imageModelDefaults["runninghub::gpt-image-2.5"].runningHubGpt25Variant, "flare");
    assert.equal(patch.imageModelDefaults["runninghub::gpt-image-2.5"].runningHubGpt25Mode, "official");
});

test("RH existing GPT Image 2 and Nano controls do not gain 2.5 modes or variants", () => {
    for (const model of ["gpt-image-2", "nano-banana-pro"]) {
        const view = render({ ...rh, model: `runninghub::${model}` });
        assert.equal(view.select(modes), undefined);
        assert.equal(view.select(variants), undefined);
        assert.equal(view.select(qualities), undefined);
        const oldQuality = view.select(["low", "medium", "high"]);
        assert.equal(Boolean(oldQuality), model === "gpt-image-2");
    }
});

test("reopening a generated node restores its RH settings instead of the current global defaults", () => {
    const metadata = {
        model: rh.model,
        runningHubGpt25Mode: "official",
        runningHubGpt25Variant: "sunburst",
        runningHubOutputFormat: "webp",
        runningHubBackground: "transparent",
        gptImageQuality: "xhigh",
        quality: "4k",
    };
    const restored = context.vke(rh, { metadata }, "image");
    for (const [key, value] of Object.entries(metadata)) assert.equal(restored[key], value, key);
    const view = render(restored);
    assert.equal(view.select(modes).props.value, "official");
    assert.equal(view.select(variants).props.value, "sunburst");
    assert.equal(view.select(qualities).props.value, "xhigh");
    assert.equal(view.select(formats).props.value, "webp");
});

test("Mart and Apilio variant callbacks remain on their preexisting independent settings", () => {
    for (const provider of ["apimart", "apilio"]) {
        const view = render({
            ...rh,
            provider,
            model: `${provider}::gpt-image-2.5`,
            apimartGpt25Variant: "flare",
        });
        view.select(variants).props.onChange("sunburst");
        assert.deepEqual(view.changes[0].patch, { apimartGpt25Variant: "sunburst" });
        if (provider === "apilio") assert.equal(view.select(modes), undefined);
    }
});

test("RH canvas offers ten fixed ratios and fifteen official ratios, each plus Auto", () => {
    for (const [mode, count] of [["fixed", 11], ["official", 16]]) {
        const view = render({ ...rh, runningHubGpt25Mode: mode });
        const ratios = view.selectors.find(node => node.props.showRatioShape);
        assert.equal(ratios.props.items.length, count);
        assert.equal(ratios.props.items.some(item => item.value === "2:1"), mode === "official");
        assert.equal(ratios.props.items[0].value, "auto");
    }
});

test("old fixed unsupported selections display Auto without rewriting the stored project", () => {
    const config = { ...rh, size: "2:1" };
    const view = render(config);
    const ratios = view.selectors.find(node => node.props.showRatioShape);
    assert.equal(ratios.props.value, "auto");
    assert.equal(config.size, "2:1");
    assert.equal(view.changes.length, 0);
    assert.equal(runningHubImageRequestSpec(config, "draw").body.aspectRatio, undefined);
});

test("switching RH Official to Fixed resets only an incompatible ratio", () => {
    const view = render({ ...rh, runningHubGpt25Mode: "official", size: "2:1" });
    view.select(modes).props.onChange("fixed");
    assert.deepEqual(view.changes[0].patch, { runningHubGpt25Mode: "fixed", size: "auto" });
    const compatible = render({ ...rh, runningHubGpt25Mode: "official", size: "16:9" });
    compatible.select(modes).props.onChange("fixed");
    assert.deepEqual(compatible.changes[0].patch, { runningHubGpt25Mode: "fixed" });
    const fixed = render({ ...rh, size: "16:9" });
    fixed.select(modes).props.onChange("official");
    assert.deepEqual(fixed.changes[0].patch, { runningHubGpt25Mode: "official" });
});

test("RH fixed ratio filtering does not remove extended ratios from other providers or RH GPT2", () => {
    for (const [provider, model] of [
        ["apilio", "gpt-image-2.5"], ["apimart", "gpt-image-2.5"],
        ["runninghub", "gpt-image-2"],
    ]) {
        const view = render({ ...rh, provider, model: `${provider}::${model}`, size: "2:1" });
        const ratios = view.selectors.find(node => node.props.showRatioShape);
        assert.equal(ratios.props.items.length, 16);
        assert.equal(ratios.props.value, "2:1");
    }
});
