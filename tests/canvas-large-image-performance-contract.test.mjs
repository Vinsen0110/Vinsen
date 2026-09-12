import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

test("canvas images remain mounted while the viewport moves", () => {
    assert.doesNotMatch(bundle, /function useCanvasImageVisibility\(\)/);
    assert.doesNotMatch(bundle, /rootMargin:"640px"/);
    assert.match(bundle, /\[d,f\]=c\.useState\(\(\)=>n\)/);
    assert.match(bundle, /if\(!n\)return f\(""\)/);
    assert.match(bundle, /children:d\?y\.jsx\("img"/);
    assert.match(bundle, /return f\(n\),acquireCanvasPreview\(/);
});

test("large previews persist in IndexedDB and are reused after reload", () => {
    assert.match(bundle, /storeName:"canvas_previews"/);
    assert.match(bundle, /function canvasPreviewPersistentKey\(/);
    assert.match(bundle, /function loadOrCreateCanvasPreview\(/);
    assert.match(bundle, /await a\.getItem\(o\)/);
    assert.match(bundle, /await a\.setItem\(o,\{blob:i,lastUsed:Date\.now\(\)\}\)/);
});

test("preview generation stays queued and both caches stay bounded", () => {
    assert.match(bundle, /t\?canvasPreviewQueue\(\)\.pending\.unshift\(o\):canvasPreviewQueue\(\)\.pending\.push\(o\)/);
    assert.match(bundle, /Date\.now\(\)-new Date\(t\.generatedAt\|\|0\)\.getTime\(\)<6e4/);
    assert.match(bundle, /for\(;e\.active<2&&e\.pending\.length;\)/);
    assert.match(bundle, /if\(e\.size<=96\)return/);
    assert.match(bundle, /setTimeout\(\(\)=>\{i\.refs===0[^}]+\},6e4\)/);
    assert.match(bundle, /t\.length>160/);
});

test("remote image caching cannot remain pending forever", () => {
    assert.match(bundle, /async function fetchImageBlobWithProgress\(e,t\)\{const n=new AbortController,r=setTimeout\(\(\)=>n\.abort\(\),6e4\)/);
    assert.match(bundle, /if\(n\.signal\.aborted\)throw new Error\("图片加载超时，请检查网络后重试"\)/);
    assert.match(bundle, /throw new Error\(`图片下载失败：\$\{o\.status\}`\)/);
    assert.match(bundle, /m&&\(g>0\|\|Number\.isFinite\(h\)&&h>0\)\?y\.jsxs\("div"/);
});

test("all generated image results are displayed before browser cache persistence", () => {
    assert.match(bundle, /function queueCanvasGeneratedImageCache\(/);
    assert.match(bundle, /function cacheCanvasGeneratedImage\(/);
    assert.equal(
        (bundle.match(/queueCanvasGeneratedImageCache\(\{source:(?:nt|Er|bn)\.dataUrl/g) || []).length,
        4,
    );
    assert.doesNotMatch(bundle, /if\(isRemoteImageUrl\((?:nt|Er|bn)\.dataUrl\)\)/);
});

test("asset re-import restores storage-backed image content and dimensions", () => {
    assert.ok(
        bundle.includes(
            'Af=c.useCallback(async O=>{const K=O.storageKey?{url:await Xl(O.storageKey,O.dataUrl||""),storageKey:O.storageKey,width:Number(O.width)||0,height:Number(O.height)||0,bytes:Number(O.bytes)||0,mimeType:O.mimeType||"image/png"}:await bo(O.dataUrl);if(!K.url)throw new Error("素材图片读取失败，请重新导入");const Q=K.width>0&&K.height>0?K:await gh(K.url)',
        ),
    );
});

test("Alt-click duplication keeps insertion in the selection update batch", () => {
    const start = bundle.indexOf("Dh=c.useCallback(");
    const end = bundle.indexOf(",Nl=c.useCallback(", start);
    assert.ok(start >= 0 && end > start);
    const duplicate = bundle.slice(start, end);
    assert.doesNotMatch(duplicate, /startTransition/);
    assert.match(duplicate, /_\(Vt=>\[\.\.\.Vt,\.\.\.Rn\]\),nt\.length&&ee\(Vt=>\[\.\.\.Vt,\.\.\.nt\]\)/);
    assert.match(duplicate, /vn\.current=Mt,ye\(gt\),oo\.current=gt/);
});
