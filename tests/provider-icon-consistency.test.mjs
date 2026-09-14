import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

function declaration(name) {
    const start = bundle.indexOf(`function ${name}(`);
    const end = bundle.indexOf("function ", start + 10);
    assert.ok(start >= 0 && end > start, `missing icon component ${name}`);
    return bundle.slice(start, end);
}

const runtime = vm.createContext({
    y: { jsx: (type, props) => ({ type, props }) },
    pr: value => value.split("::").at(-1),
    p7: "fallback-model-icon",
    Is: "fallback-provider-icon",
});
vm.runInContext(
    `${declaration("nq")}${declaration("Q2e")}${declaration("yX")}`,
    runtime,
);

function modelIcon(model) {
    return runtime.nq({ model });
}

function providerIcon(provider) {
    return runtime.yX({ provider });
}

test("Google canvas and settings icons use the same unfiltered local artwork", () => {
    const canvas = providerIcon("Google");
    for (const model of ["nano-banana-pro", "gemini-3.8-flash", "apimart::nano-banana-pro"]) {
        const settings = modelIcon(model);
        assert.equal(settings.props.src, "./icons/gemini.svg");
        assert.equal(settings.props.src, canvas.props.src);
        assert.equal(settings.props.className, canvas.props.className);
        assert.doesNotMatch(settings.props.className, /invert|brightness|saturate|filter/);
    }
});

test("OpenAI canvas and settings icons keep the same dark-mode white treatment", () => {
    const canvas = providerIcon("OpenAI");
    for (const model of ["gpt-image-2", "gpt-image-2.5"]) {
        const settings = modelIcon(model);
        assert.equal(settings.props.src, "./icons/openai.svg");
        assert.equal(settings.props.className, canvas.props.className);
        assert.match(settings.props.className, /dark:invert/);
    }
});

test("provider icons reserve a stable nonshrinking square in both selectors", () => {
    for (const icon of [
        providerIcon("Google"), providerIcon("OpenAI"),
        modelIcon("nano-banana-pro"), modelIcon("gpt-image-2.5"),
    ]) {
        assert.equal(icon.props.width, 16);
        assert.equal(icon.props.height, 16);
        for (const token of ["canvas-provider-icon", "size-4", "shrink-0", "object-contain"]) {
            assert.ok(icon.props.className.split(" ").includes(token));
        }
    }
});

test("unrelated providers preserve their existing icon and fallback behavior", () => {
    const claude = modelIcon("claude");
    assert.equal(claude.props.src, "./icons/claude.svg");
    assert.match(claude.props.className, /dark:brightness-\[1\.3\] dark:saturate-\[1\.1\]/);
    assert.equal(modelIcon("unrecognized-model").type, "fallback-model-icon");
    assert.equal(providerIcon("Other").type, "fallback-provider-icon");
});
