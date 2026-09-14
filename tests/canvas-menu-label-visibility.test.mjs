import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const html = await readFile(new URL("../index.html", import.meta.url), "utf8");
const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const api = await readFile(new URL("../apimart-api.js", import.meta.url), "utf8");

test("the dark brand column overrides the exact light-mode container and selected buttons", () => {
    assert.match(html, /\.dark \.canvas-model-menu > div > div > div:first-child \{[^}]*background: #171717 !important;/);
    assert.match(html, /\.dark \.canvas-model-menu > div > div > div:first-child > button\.bg-white \{[^}]*background: #292524 !important;/);
    assert.match(html, /\.dark \.canvas-model-menu > div > div > div:first-child > button \{[^}]*color: #d6d3d1 !important;/);
});

test("parameter triggers size to the complete label instead of their old fixed width", () => {
    assert.match(html, /\.canvas-generation-select \{[^}]*width: max-content !important;[^}]*min-width: max-content !important;/);
    assert.match(html, /\.canvas-generation-select > span \{[^}]*flex: 0 0 auto;[^}]*overflow: visible;[^}]*text-overflow: clip;[^}]*white-space: nowrap;/);
    assert.match(html, /\.canvas-generation-select > svg \{[^}]*flex-shrink: 0;/);
    assert.match(html, /controlsWidth = items\.reduce[\s\S]*?item\.offsetWidth/);
});

test("constrained toolbars scroll their intact controls without covering the generate action", () => {
    assert.match(html, /\.canvas-generation-toolbar > div:first-child \{[^}]*min-width: 0;[^}]*overflow-x: auto;[^}]*overflow-y: hidden;/);
    assert.match(html, /\.canvas-generation-panel:has\(\.canvas-model-select\[title\^="gpt-image-2"\]\) \.canvas-generation-toolbar > div:first-child \{[^}]*overflow-x: auto;[^}]*overflow-y: hidden;/);
});

test("Google retains brand colors while OpenAI still turns white in dark mode", () => {
    assert.match(html, /img\[src\*="\/gemini\.svg"\] \{[^}]*filter: none !important;/);
    assert.doesNotMatch(html, /\.dark img\[src\*="\/gemini\.svg"\]/);
    assert.match(html, /\.dark img\[src\*="\/openai\.svg"\],[^{]+\{[^}]*filter: invert\(1\) !important;/);
});

test("withdrawn Seedream model and dependency are absent from shipped code", () => {
    for (const source of [html, bundle, api]) {
        assert.doesNotMatch(source, /seedream-5-0-pro|seedream-canvas|volcengine\.svg/);
    }
});
