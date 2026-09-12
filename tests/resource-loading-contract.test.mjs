import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");
const tags = [...indexHtml.matchAll(/<(link|script|style)\b([^>]*)>/gi)].map((match) => {
    const attributes = new Map(
        [...match[2].matchAll(/([^\s=/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g)]
            .map((attribute) => [
                attribute[1].toLowerCase(),
                attribute[2] ?? attribute[3] ?? attribute[4] ?? "",
            ]),
    );
    return { name: match[1].toLowerCase(), attributes, index: match.index };
});

const entryScripts = tags.filter(({ name, attributes }) =>
    name === "script"
    && attributes.get("type") === "module"
    && /^\.\/assets\/index-[^/?]+\.js(?:\?|$)/.test(attributes.get("src") || ""),
);
const stylesheets = tags.filter(({ name, attributes }) =>
    name === "link"
    && attributes.get("rel") === "stylesheet"
    && /^\.\/assets\/index-[^/?]+\.css(?:\?|$)/.test(attributes.get("href") || ""),
);

function assertSameCrossorigin(preload, resource) {
    assert.equal(
        preload.attributes.has("crossorigin"),
        resource.attributes.has("crossorigin"),
        "preload and resource must agree on crossorigin presence",
    );
    assert.equal(
        preload.attributes.get("crossorigin"),
        resource.attributes.get("crossorigin"),
        "preload and resource must use the same crossorigin value",
    );
}

test("the entry module preload matches the actual script URL and crossorigin", () => {
    assert.equal(entryScripts.length, 1);
    const entry = entryScripts[0];
    const preloads = tags.filter(({ name, attributes }) =>
        name === "link"
        && attributes.get("rel") === "modulepreload"
        && attributes.get("href") === entry.attributes.get("src"),
    );
    assert.equal(preloads.length, 1, "the exact entry URL needs one modulepreload");
    assertSameCrossorigin(preloads[0], entry);
    assert.ok(preloads[0].index < entry.index);
});

test("the entry bundle has no legacy as-script preload", () => {
    const legacyPreloads = tags.filter(({ name, attributes }) =>
        name === "link"
        && attributes.get("rel") === "preload"
        && attributes.get("as") === "script",
    );
    assert.equal(legacyPreloads.length, 0, "ES modules must use modulepreload, not as=script");
});

test("the stylesheet preload matches its original URL, type, and crossorigin", () => {
    assert.equal(stylesheets.length, 1);
    const stylesheet = stylesheets[0];
    const preloads = tags.filter(({ name, attributes }) =>
        name === "link"
        && attributes.get("rel") === "preload"
        && attributes.get("href") === stylesheet.attributes.get("href"),
    );
    assert.equal(preloads.length, 1, "the exact stylesheet URL needs one preload");
    assert.equal(preloads[0].attributes.get("as"), "style");
    assertSameCrossorigin(preloads[0], stylesheet);
    assert.ok(preloads[0].index < stylesheet.index);
});

test("the original stylesheet stays by the entry and before inline overrides", () => {
    assert.equal(entryScripts.length, 1);
    assert.equal(stylesheets.length, 1);
    const entry = entryScripts[0];
    const stylesheet = stylesheets[0];
    const inlineStyles = tags.filter(({ name }) => name === "style");
    assert.ok(inlineStyles.length > 0);
    assert.ok(stylesheet.index > entry.index, "do not move the stylesheet to the preload slot");
    assert.ok(
        inlineStyles.every(({ index }) => index > stylesheet.index),
        "inline style overrides must remain after the original stylesheet",
    );
    assert.ok(
        indexHtml.slice(entry.index, inlineStyles[0].index).includes(
            `href="${stylesheet.attributes.get("href")}"`,
        ),
        "the real stylesheet must remain near the module entry",
    );
});
