import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");

test("image nodes keep their natural ratio while other canvas nodes resize freely", () => {
    assert.doesNotMatch(bundle, /Ye=220,it=160/);
    assert.match(bundle, /Ye=64,it=48/);
    assert.match(bundle, /Dt=Math\.max\(Ye,Ce\.current\.startWidth/);
    assert.match(bundle, /dt=Math\.max\(it,Ce\.current\.startHeight/);
    assert.match(bundle, /Pt=e\.type===Ne\.Image&&!!e\.metadata\?\.content/);
    assert.match(bundle, /Number\(e\.metadata\?\.naturalWidth\)\|\|Ce\.current\.startWidth/);
    assert.match(bundle, /Number\(e\.metadata\?\.naturalHeight\)\|\|Ce\.current\.startHeight/);
    assert.match(bundle, /ke=Math\.max\(Ye,it\*Vt/);
    assert.match(bundle, /qe=ke\/Vt/);
});

test("text editors no longer carry hard min/max resize bounds", () => {
    assert.match(bundle, /function canvasTextBoxResizeStyle\(e,t\)\{return\{width:e,height:t,resize:"none"/);
    assert.doesNotMatch(bundle, /minWidth:280,minHeight:80,maxWidth:900,maxHeight:600/);
    assert.doesNotMatch(bundle, /resize:"both"/);
    assert.match(indexHtml, /\.canvas-generation-panel:not\(\[data-canvas-resize-ready="true"\]\) textarea,[\s\S]*?height: auto !important;/s);
    assert.match(indexHtml, /\.canvas-generation-panel > \.relative\.rounded-xl\.border > \[contenteditable="true"\] \{[^}]*flex-basis: 132px;/s);
});

test("text generation panels use the same outer-frame resize contract", () => {
    assert.match(bundle, /function lke\([\s\S]*?className:"canvas-generation-panel/);
    assert.doesNotMatch(bundle, /function lke\([\s\S]*?canvasTextBoxResizeStyle\(596,112\)/);
    assert.match(indexHtml, /\.canvas-generation-panel > \.p-3 \{[^}]*display: flex;[^}]*flex: 1 1 auto;[^}]*min-height: 0;/s);
    assert.match(indexHtml, /\.canvas-generation-panel\[data-canvas-resize-ready="true"\] textarea,[\s\S]*?width: 100% !important;[^}]*min-width: 0 !important;[^}]*min-height: 0 !important;[^}]*max-width: none !important;/s);
});

test("generation panels derive their default width from visible controls before a user resize", () => {
    assert.match(
        indexHtml,
        /\.canvas-generation-panel \{[\s\S]*?width: min\(var\(--canvas-panel-auto-width, 660px\), calc\(100vw - 32px\)\) !important;/s,
    );
    assert.match(indexHtml, /var updateAutoWidth = function \(panel\)/);
    assert.match(bundle, /canvasTextBoxResizeStyle\(636,96\)/);
    assert.match(bundle, /canvasTextBoxResizeStyle\(636,132\)/);
});
