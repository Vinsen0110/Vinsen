import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");

test("both existing canvas create menus expose text generation", () => {
    const connectionMenu = bundle.slice(bundle.indexOf("function zke("), bundle.indexOf("function eB("));
    const quickMenu = bundle.slice(bundle.indexOf("function kke("), bundle.indexOf("function G0("));

    assert.match(connectionMenu, /title:"\\u6587\\u672C\\u751F\\u6210"/);
    assert.match(connectionMenu, /onClick:\(\)=>t\("text-generation"\)/);
    assert.match(quickMenu, /label:"\\u6587\\u672C\\u751F\\u6210"/);
    assert.match(quickMenu, /onClick:\(\)=>t\("text-generation"\)/);
    assert.doesNotMatch(connectionMenu, /audio-input/);
    assert.doesNotMatch(quickMenu, /audio-input/);
});

test("double-click text generation creates a configured text node", () => {
    assert.match(
        bundle,
        /O==="text-generation"\)\{Lu\(Ne\.Text,"\\u6587\\u672C\\u751F\\u6210",K,\{content:"",prompt:"",status:rd,fontSize:14,model:ODe\(N,"text",N\.textModel\)\}\)/,
    );
});

test("connected text generation becomes reverse prompt for an image source", () => {
    assert.match(bundle, /isReverseText=O===Ne\.Text&&sourceNode\?\.type===Ne\.Image/);
    assert.match(bundle, /\.\.\.isReverseText\?\{reversePrompt:!0[^}]*\}:\{\}/);
    assert.match(
        bundle,
        /title:isReverseText\?"\\u53CD\\u63A8\\u63D0\\u793A\\u8BCD":"\\u6587\\u672C\\u751F\\u6210"/,
    );
    assert.match(
        bundle,
        /O==="text-generation"\?c1\(Ne\.Text,Ge\):c1\(Ne\.Image,Ge\)/,
    );
    assert.match(bundle, /O!==Ne\.Audio&&ln\(le\.id\)/);
    assert.equal(
        bundle.match(/model:ODe\(N,"text",N\.textModel\)/g)?.length,
        3,
        "all canvas text-node creation paths should resolve the current site's text model",
    );
});

test("text-generation nodes do not expose the image-generation action", () => {
    assert.match(
        bundle,
        /\.\.\.fe&&!fX\(e\)&&e\.title!=="\\u6587\\u672C\\u751F\\u6210"\?\[\{id:"generateImage"/,
    );
    assert.match(
        bundle,
        /children:\[!fX\(e\)&&e\.title!=="\\u6587\\u672C\\u751F\\u6210"\?y\.jsxs\("button",\{type:"button",className:"absolute right-3 top-3/,
        "the inline image-generation button must also be hidden",
    );
});

test("reverse-prompt references remain connected and visible", () => {
    assert.match(
        bundle,
        /\.map\(st=>\(\{id:cn\(\),\.\.\.st\}\)\)/,
        "dragging from an image should create a normal visible connection",
    );
    assert.match(
        bundle,
        /fromNodeId:O\.id,toNodeId:Ee\.id\}/,
        "the toolbar reverse-prompt action should create a visible connection",
    );
    assert.match(bundle, /return!!\(K&&Q&&!aB\(K,B\)&&!aB\(Q,B\)/);
    assert.doesNotMatch(bundle, /!O\.hidden&&K&&Q/);
    assert.doesNotMatch(bundle, /!fX\(Q\)&&!aB\(K,B\)/);
});

test("reverse-prompt result cards do not render an internal divider", () => {
    assert.match(indexHtml, /div:has\(> \[aria-label="反推分析结果"\]\) > \.border-b/);
    assert.match(indexHtml, /border:\s*0/);
    assert.match(indexHtml, /position:\s*absolute/);
});

test("connected text can be used directly as an image prompt", () => {
    assert.match(
        bundle,
        /D=v==="image"\?i\.filter\(z=>z\.active&&z\.kind==="text"&&String\(z\.text\|\|""\)\.trim\(\)\):\[\],B=D\.length>0/,
    );
    assert.match(bundle, /\(!z&&!A&&!B\)\|\|t\|\|o\(e\.id,v,z\)/);
    assert.match(bundle, /canSubmit:A\|\|B,credits:j/);
    assert.match(bundle, /function V6\([\s\S]+?prompt:i\?`\$\{r\}[\s\S]*?\$\{i\}`:r/);
});

test("the reverse-prompt preset requests structured bilingual visual detail", () => {
    const match = bundle.match(/const W0=\x60([\s\S]*?)\x60;/);
    assert.ok(match, "the shared reverse-prompt preset should exist");
    const preset = JSON.parse('"' + match[1].replaceAll("\n", "\\n") + '"');

    for (const section of [
        "主体与造型",
        "构图与机位",
        "材质与工艺",
        "光源与阴影",
        "场景与空间",
        "【中文提示词】",
        "【English Prompt】",
        "不要输出分析步骤",
    ]) {
        assert.match(preset, new RegExp(section));
    }
    assert.match(bundle, /,bg=W0;function sanitizeCanvasNodeClone/);
});

test("connected image and text references are visible in generation panels", () => {
    const panel = bundle.slice(bundle.indexOf("function ske("), bundle.indexOf("function fke("));

    assert.match(
        panel,
        /T=\(v==="image"\|\|v==="text"\)\?i\.filter\(z=>z\.active&&z\.kind==="image"&&z\.previewUrl\):\[\]/,
        "text panels should keep active image references for their thumbnail strip",
    );
    assert.match(panel, /references:T,nodeId:e\.id,\s*onReferenceOrderChange:handleReferenceOrder/);
    assert.match(panel, /handleReferenceOrder=\(nodeId,ids\)=>\{\$\(remapImageMentions\(P,T\.map\(item=>item\.nodeId\),ids\)\);l\?\.\(nodeId,ids\)\}/);
    assert.match(panel, /referenceImages\.length\?y\.jsx\(dke,\{nodeId:referenceNodeId/);
    assert.match(
        panel,
        /D=v==="image"\?i\.filter\(z=>z\.active&&z\.kind==="text"&&String\(z\.text\|\|""\)\.trim\(\)\):\[\]/,
        "image panels should collect active upstream text",
    );
    assert.match(panel, /D\.length\?y\.jsx\(CanvasTextReferenceStrip/);
    assert.match(panel, /data-canvas-text-references/);
    assert.match(panel, /canvas-text-reference-preview/);
    assert.match(indexHtml, /\.canvas-text-reference-strip \{/);
    assert.match(indexHtml, /\.dark \.canvas-text-reference-item \{/);
});

test("image generation panels keep a stable width when text references are long", () => {
    const panel = bundle.slice(bundle.indexOf("function ske("), bundle.indexOf("function fke("));

    assert.match(panel, /style:\{width:v==="image"\?660:void 0,maxWidth:v==="image"\?660:void 0/);
    assert.match(panel, /v==="image"\?\{width:"100%",maxWidth:"100%",resize:"none"\}:\{\}/);
});

test("the connection resolver passes generated text into an empty image node", () => {
    const extractFunction = (name) => {
        const start = bundle.indexOf(`function ${name}(`);
        assert.ok(start >= 0, `${name} should exist`);
        const end = bundle.indexOf("function ", start + 10);
        return bundle.slice(start, end);
    };
    const resolvePrompt = new Function(
        "Ne",
        ["uze", "pI", "rX", "gI", "oX", "V6", "aX", "mze", "pze", "gze", "yze"]
            .map(extractFunction)
            .join("\n") + "\nreturn V6;",
    )({ Image: "image", Video: "video", Audio: "audio", Text: "text", Config: "config" });

    const result = resolvePrompt(
        "image-node",
        [
            { id: "text-node", type: "text", metadata: { content: "cinematic sunset over the sea" } },
            { id: "image-node", type: "image", metadata: {} },
        ],
        [{ fromNodeId: "text-node", toNodeId: "image-node" }],
        "",
    );

    assert.equal(result.prompt.trim(), "cinematic sunset over the sea");
    assert.equal(result.textCount, 1);
});
