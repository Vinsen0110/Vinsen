import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const html = await readFile(new URL("../index.html", import.meta.url), "utf8");
const viewer = bundle.slice(bundle.indexOf("function Bke("), bundle.indexOf("function SX("));

test("image viewer scopes its overlay, metadata and controls to the current theme", () => {
    for (const name of ["canvas-image-viewer", "canvas-image-viewer-control", "canvas-image-viewer-info", "canvas-image-viewer-title", "canvas-image-viewer-image"]) {
        assert.ok(viewer.includes(name), name);
    }
    assert.doesNotMatch(viewer, /bg-white|text-stone-|border-stone-/);
    assert.match(html, /\.dark \.canvas-image-viewer\s*\{[^}]*--viewer-overlay:\s*rgba\(15, 16, 18, \.92\)/);
    for (const token of ["panel", "hover", "border", "text", "muted", "title", "shadow"]) {
        assert.match(html, new RegExp(`\\.dark \\.canvas-image-viewer\\s*\\{[^}]*--viewer-${token}:`));
    }
});

test("viewer styling never recolors the image and keeps mobile info above the zoom control", () => {
    const imageRule = html.match(/\.canvas-image-viewer-image\s*\{([^}]+)\}/)?.[1];
    assert.ok(imageRule);
    assert.doesNotMatch(imageRule, /(?:filter|opacity|mix-blend-mode)\s*:/);
    assert.match(html, /@media \(max-width: 767px\)\s*\{\s*\.canvas-image-viewer-info\s*\{[^}]*bottom: 72px;/);
});

test("preview keeps zoom, drag, Escape, original-image download and propagation guards", () => {
    assert.match(viewer, /Math\.min\(8,Math\.max\(\.25,n\*/);
    assert.match(viewer, /P\.key==="Escape"/);
    assert.match(viewer, /setPointerCapture\(S\.pointerId\)/);
    assert.match(viewer, /r\(1\),a\(\{x:0,y:0\}\)/);
    assert.match(viewer, /await fetch\(f\),T=await j\.blob\(\)/);
    assert.match(viewer, /await N\.write\(T\),await N\.close\(\)/);
    assert.match(viewer, /onPointerDown:S=>S\.stopPropagation\(\)/);
});
