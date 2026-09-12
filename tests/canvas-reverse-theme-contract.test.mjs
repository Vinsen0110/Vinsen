import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");

function section(startToken, endToken) {
    const start = bundle.indexOf(startToken);
    const end = bundle.indexOf(endToken, start);
    assert.ok(start >= 0 && end > start, `missing source section: ${startToken}`);
    return bundle.slice(start, end);
}

const shell = section("Bze=xe.memo(", "function Fze(");
const text = section("function Gze(", "function Kze(");
const result = section("function Kze(", "function Xze(");

test("reverse theme scope is applied only to reverse-prompt nodes", () => {
    assert.match(
        bundle,
        /function fX\(e\)\{return e\.type===Ne\.Text&&\(!!e\.metadata\?\.reversePrompt\|\|e\.title==="\\u53CD\\u63A8\\u63D0\\u793A\\u8BCD"\)\}/,
    );
    assert.match(shell, /oe=fX\(e\)/);
    assert.ok(shell.includes('${oe?"canvas-reverse-shell":""}'));
    assert.match(text, /fX\(e\)&&!n\?y\.jsx\(Kze,\{node:e\}\)/);
    assert.match(text, /color:t\.node\.text/);
    assert.match(text, /"bg-transparent pl-4 pr-14 pt-0 pb-4 font-sans"/);
});

test("both reverse frame layers use variables while selected borders stay green", () => {
    assert.equal((shell.match(/background:oe\?"var\(--canvas-reverse-frame\)"/g) || []).length, 2);
    assert.match(shell, /ge="#22c55e"/);
    assert.match(shell, /borderColor:oe\?Pe\?ge:"var\(--canvas-reverse-border\)"/);
    assert.doesNotMatch(shell, /background:oe\?"rgba\(255,255,255/);
    assert.match(shell, /oe\?null:y\.jsx\(Yze,\{node:e\}\)/);
});

test("reverse card light and dark theme colors are explicitly scoped", () => {
    const light = indexHtml.match(/(?:^|\n)\s*\.canvas-reverse-shell\s*\{([^}]+)\}/)?.[1];
    const dark = indexHtml.match(/\.dark \.canvas-reverse-shell\s*\{([^}]+)\}/)?.[1];
    assert.ok(light);
    assert.ok(dark);
    assert.match(light, /--canvas-reverse-frame:\s*rgba\(255,\s*255,\s*255,\s*\.96\)/);
    assert.match(light, /--canvas-reverse-surface:\s*#fff;/);
    assert.match(light, /--canvas-reverse-text:\s*#57534e;/);
    assert.match(light, /--canvas-reverse-border:\s*rgba\(226,\s*232,\s*240,\s*\.95\)/);
    assert.match(dark, /--canvas-reverse-frame:\s*#292524;/);
    assert.match(dark, /--canvas-reverse-surface:\s*#292524;/);
    assert.match(dark, /--canvas-reverse-text:\s*#f5f5f4;/);
    assert.match(dark, /--canvas-reverse-border:\s*#44403c;/);
    assert.match(
        indexHtml,
        /\.canvas-reverse-card,\s*\.canvas-reverse-content\s*\{[^}]*color:\s*var\(--canvas-reverse-text\);[^}]*background:\s*var\(--canvas-reverse-surface\);/,
    );
});

test("reverse display and edit content share padding, type size, and line height", () => {
    const editClass = text.match(/fX\(e\)\?"(canvas-reverse-content[^"]+)"/)?.[1] || "";
    const displayClass = result.match(/className:"(canvas-reverse-content[^"]+)"/)?.[1] || "";
    for (const className of [editClass, displayClass]) {
        const tokens = new Set(className.split(/\s+/));
        for (const token of ["canvas-reverse-content", "px-4", "py-3", "text-[14px]", "leading-7"]) {
            assert.ok(tokens.has(token), `missing shared reverse content class: ${token}`);
        }
        assert.ok(!tokens.has("bg-white"));
        assert.ok(!tokens.has("text-stone-700"));
        assert.ok(!tokens.has("pt-12"));
    }
    assert.match(result, /className:"canvas-reverse-card /);
    assert.doesNotMatch(result, /\bbg-white\b|\btext-stone-700\b/);
});

test("reverse editor has no hardcoded inline color overriding its theme", () => {
    const editorStyle = text.match(/style:fX\(e\)\?\{([^}]+)\}:f/)?.[1] || "";
    assert.ok(editorStyle);
    assert.match(editorStyle, /fontSize:"14px"/);
    assert.match(editorStyle, /lineHeight:"28px"/);
    assert.doesNotMatch(editorStyle, /(?:^|,)(?:color|background|backgroundColor):/);
});

test("reverse double-click editing, blur, Escape, and propagation guards remain intact", () => {
    assert.match(shell, /e\.type===Ne\.Text&&\(_e\.stopPropagation\(\),B\(!0\)\)/);
    assert.match(text, /onChange:m=>a\(e\.id,m\),onBlur:i,onKeyDown:m=>\{m\.key==="Escape"&&i\(\)\}/);
    assert.match(
        text,
        /onMouseDown:m=>m\.stopPropagation\(\),onPointerDown:m=>m\.stopPropagation\(\),onWheel:m=>m\.stopPropagation\(\)/,
    );
    assert.match(result, /onWheel:r=>r\.stopPropagation\(\)/);
});

test("reverse copy controls keep their clipboard action and use the shared text color", () => {
    assert.match(result, /onClick:r=>\{r\.stopPropagation\(\),n\(t,"\\u53CD\\u63A8\\u6587\\u5B57\\u5DF2\\u590D\\u5236"\)\}/);
    assert.match(result, /onMouseDown:r=>r\.stopPropagation\(\),onPointerDown:r=>r\.stopPropagation\(\)/);
    assert.match(result, /"aria-label":"\\u590D\\u5236\\u53CD\\u63A8\\u6587\\u5B57"/);
    assert.match(
        indexHtml,
        /\.canvas-reverse-shell button\[aria-label="复制反推文字"\]\s*\{[^}]*color:\s*var\(--canvas-reverse-text\);/,
    );
    assert.match(
        indexHtml,
        /\[data-node-id\] div:has\(> \[aria-label="反推分析结果"\]\) > \.border-b\s*\{[^}]*border:\s*0;/,
    );
});
