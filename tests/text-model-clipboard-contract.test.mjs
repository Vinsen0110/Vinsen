import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { APIMART_TEXT_MODELS } from "../apimart-api.js";
import tudouProxy from "../api/tudou-proxy.js";
import { usesCloudinaryReferenceHost } from "../cloudinary-reference-upload.js";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

test("supported text models stay isolated by site", () => {
    assert.match(bundle, /APOLLO_TEXT_MODELS=\["gemini-3\.8-flash"\]/);
    assert.match(bundle, /TUDOU_TEXT_MODELS=\["gpt-5\.5"\]/);
    assert.match(
        bundle,
        /function siteTextModelNames\(e\)\{return e===RUNNINGHUB_SITE_ID\?RUNNINGHUB_TEXT_MODELS:e===TUDOU_SITE_ID\?TUDOU_TEXT_MODELS:e===GRSAI_SITE_ID\?GRSAI_TEXT_MODELS:e===APIMART_SITE_ID\?APIMART_TEXT_MODELS:APOLLO_TEXT_MODELS\}/,
    );
    assert.deepEqual(APIMART_TEXT_MODELS, []);
    assert.match(bundle, /textModels:siteModelRefs\(t,siteTextModelNames\(t\)\)/);
    assert.match(
        bundle,
        /const f=Array\.from\(new Set\(e\.textModels\|\|\[\]\)\),m=f\.includes\(e\.model\)/,
    );
    assert.doesNotMatch(bundle, /gpt-5\.6-sol/);
    assert.doesNotMatch(bundle, /filter\(g=>!pr\(g\)\.toLowerCase\(\)\.includes\("gemini"\)\)/);
});

test("text model labels are unified without changing site request IDs", () => {
    assert.match(
        bundle,
        /const APILIO_TEXT_MODEL_NAME="gemini-3\.8-flash",UNIFIED_TEXT_MODEL_NAME="gemini-3\.7-flash";/,
    );
    assert.match(
        bundle,
        /return r===DP&&n===APILIO_TEXT_MODEL_NAME\?APILIO_TEXT_MODEL_NAME:n===UNIFIED_TEXT_MODEL_NAME\|\|siteTextModelNames\(r\)\.includes\(n\)\?UNIFIED_TEXT_MODEL_NAME:n/,
    );
    assert.match(bundle, /r==="text"\?displayTextModelName\(e,v\):pr\(v\)/);
    assert.match(bundle, /textValue:r==="text"\?displayTextModelName\(e,w\):pr\(w\)/);
    assert.match(bundle, /APOLLO_TEXT_MODELS=\["gemini-3\.8-flash"\]/);
    assert.match(bundle, /APOLLO_SITE_MODELS=\[[^\]]*gemini-3\.8-flash/);
    assert.doesNotMatch(bundle, /APOLLO_TEXT_MODELS=\["gemini-3\.7-flash"\]/);
    assert.match(bundle, /l7=\["default::gemini-3\.8-flash"\]/);
    assert.match(bundle, /textModel:"default::gemini-3\.8-flash"/);
    assert.doesNotMatch(bundle, /APOLLO_TEXT_MODELS=\["gemini-3\.6-flash"\]/);
    assert.match(bundle, /TUDOU_TEXT_MODELS=\["gpt-5\.5"\]/);
    assert.match(bundle, /RUNNINGHUB_TEXT_MODELS/);
});

test("settings expose an independent global text site and model", () => {
    assert.match(bundle, /className:"generation-settings-sections"/);
    assert.match(bundle, /children:"\\u751F\\u56FE\\u6A21\\u578B\\u8BBE\\u7F6E"/);
    assert.match(bundle, /children:"\\u6587\\u672C\\u6A21\\u578B\\u8BBE\\u7F6E"/);
    assert.match(bundle, /label:"\\u9ED8\\u8BA4\\u6587\\u672C\\u8BF7\\u6C42\\u7AD9\\u70B9"/);
    assert.match(bundle, /label:"\\u9ED8\\u8BA4\\u6587\\u672C\\u6A21\\u578B（\\u5168\\u5C40）"/);
    assert.match(
        bundle,
        /value:a\.textModel\|\|defaultTextModelRef\(a\.textSiteId\|\|DP\),capability:"text"/,
    );
    assert.match(bundle, /onChange:\$=>h\(\{textModel:\$\}\)/);
});

test("Apilio, Tudou, and RH text generation use Chat Completions messages", async () => {
    const start = bundle.indexOf("function chatCompletionMessages");
    const end = bundle.indexOf("function SA", start);
    assert.ok(start >= 0 && end > start, "Chat Completions adapter should exist");

    const adapter = bundle.slice(start, end);
    const requests = [];
    const makeRequest = new Function(
        "fetch",
        "Gy",
        "xA",
        "isTudouSite",
        "isRunningHubSite",
        "RUNNINGHUB_LLM_ORIGIN",
        "wA",
        "aW",
        "AL",
        "prepareTextChatMessages",
        "assertTudouWebRequestSize",
        `${adapter};return requestChatCompletions;`,
    );
    const request = makeRequest(
        async (url, init) => {
            requests.push({ url, init });
            return new Response(JSON.stringify({
                choices: [{ message: { content: "反推结果" } }],
            }), { status: 200, headers: { "Content-Type": "application/json" } });
        },
        (baseUrl, path) => `${baseUrl.replace(/\/$/, "")}/v1${path}`,
        (_config, path) => `https://www.vinsen.top/api/tudou/v1${path}`,
        (config) => config.provider === "tudou",
        (config) => config.provider === "runninghub",
        "https://llm.runninghub.ai",
        () => ({ Authorization: "Bearer test-key", "Content-Type": "application/json" }),
        async () => "request failed",
        (payload) => {
            if (payload?.error?.message) throw new Error(payload.error.message);
        },
        async (messages) => messages,
        () => {},
    );

    const progress = [];
    const result = await request(
        { baseUrl: "https://api.apilio.ai", apiKey: "test-key", model: "gemini-3.8-flash" },
        [{
            role: "user",
            content: [
                { type: "text", text: "反推这张图" },
                { type: "image_url", image_url: { url: "data:image/png;base64,dGVzdA==" } },
            ],
        }],
        (value) => progress.push(value),
    );

    assert.equal(requests.length, 1);
    assert.equal(requests[0].url, "https://api.apilio.ai/v1/chat/completions");
    const body = JSON.parse(requests[0].init.body);
    assert.equal(body.model, "gemini-3.8-flash");
    assert.deepEqual(body.messages, [{
        role: "user",
        content: [
            { type: "text", text: "反推这张图" },
            { type: "image_url", image_url: { url: "data:image/png;base64,dGVzdA==" } },
        ],
    }]);
    assert.equal(body.stream, false);
    assert.equal("input" in body, false);
    assert.equal(result.content, "反推结果");
    assert.deepEqual(progress, ["反推结果"]);

    await request(
        {
            provider: "tudou",
            baseUrl: "https://api.ai-tudou.net",
            apiKey: "test-key",
            model: "gpt-5.5",
        },
        [{ role: "user", content: "生成一段文字" }],
    );
    assert.equal(requests.length, 2);
    assert.equal(requests[1].url, "https://www.vinsen.top/api/tudou/v1/chat/completions");
    const tudouBody = JSON.parse(requests[1].init.body);
    assert.equal(tudouBody.model, "gpt-5.5");
    assert.deepEqual(tudouBody.messages, [{ role: "user", content: "生成一段文字" }]);
    assert.equal(tudouBody.stream, false);
    assert.equal(requests[1].init.headers.Accept, "application/json");

    await request(
        {
            provider: "runninghub",
            baseUrl: "https://www.runninghub.ai",
            apiKey: "test-key",
            model: "google/gemini-3.7-flash",
        },
        [{ role: "user", content: "生成一段文字" }],
    );
    assert.equal(requests.length, 3);
    assert.equal(requests[2].url, "https://llm.runninghub.ai/v1/chat/completions");
    const runningHubBody = JSON.parse(requests[2].init.body);
    assert.equal(runningHubBody.model, "google/gemini-3.7-flash");
    assert.deepEqual(runningHubBody.messages, [{ role: "user", content: "生成一段文字" }]);
    assert.equal(runningHubBody.stream, false);
});

test("text routing has no Responses endpoint", () => {
    assert.match(bundle, /await requestChatCompletions\(o,tW\(o,t\),n,r\)/);
    assert.match(bundle, /o\.apiFormat==="gemini"&&!isTudouSite\(o\)/);
    assert.match(bundle, /isTudouSite\(e\)\?xA\(e,"\/chat\/completions"\):isRunningHubSite\(e\)\?Gy\(RUNNINGHUB_LLM_ORIGIN,"\/chat\/completions"\):Gy\(e\.baseUrl,"\/chat\/completions"\)/);
    assert.doesNotMatch(bundle, /if\(isRunningHubSite\(o\)\)throw new Error/);
    assert.doesNotMatch(bundle, /if\(isRunningHubSite\(i\)\)throw new Error/);
    assert.doesNotMatch(bundle, /xA\(e,"\/responses"\)/);
    assert.doesNotMatch(bundle, /fetch\([^)]*"\/responses"/);
});

test("large text reference images are compressed only in the request copy", async () => {
    const start = bundle.indexOf("const TEXT_REFERENCE_MAX_EDGE");
    assert.ok(start >= 0, "temporary text reference compressor should exist");
    const compressor = bundle.slice(start, bundle.indexOf("assertTudouWebRequestSize=", start));
    let disposed = false;
    let fetchedBlob = new Blob([new Uint8Array(8 * 1024 * 1024)], { type: "image/png" });
    const encodeCalls = [];
    const uploads = [];
    const makeCompressor = new Function(
        "fetch",
        "P$e",
        "throwIfTudouReferenceAborted",
        "decodeTudouReferenceBlob",
        "encodeTudouReferenceWebp",
        "Ln",
        "isImgBbReferenceUrl",
        "isTudouSite",
        "isApilioSite",
        "isRunningHubSite",
        "usesCloudinaryReferenceHost",
        "cloudinaryReferenceSource",
        `${compressor};return prepareTextChatMessages;`,
    );
    const compressMessages = makeCompressor(
        async () => ({
            ok: true,
            blob: async () => fetchedBlob,
        }),
        async (blob) => `data:image/webp;base64,${"A".repeat(Math.ceil(blob.size * 4 / 3))}`,
        (signal) => {
            if (signal?.aborted) throw new Error("aborted");
        },
        async () => ({
            image: {},
            width: 4096,
            height: 3072,
            dispose: () => { disposed = true; },
        }),
        async (_image, width, height, quality) => {
            encodeCalls.push({ width, height, quality });
            const size = Math.ceil(width * height * quality * 0.48);
            return new Blob([new Uint8Array(size)], { type: "image/webp" });
        },
        {
            post: async (url, form) => {
                uploads.push({ url, form });
                return {
                    data: {
                        success: true,
                        data: { url: "https://i.ibb.co/example/text-reference.webp" },
                    },
                };
            },
        },
        (url) => String(url).startsWith("https://i.ibb.co/"),
        (config) => config?.provider === "tudou",
        (config) => config?.provider === "apilio",
        (config) => config?.provider === "runninghub",
        usesCloudinaryReferenceHost,
        async () => "https://res.cloudinary.com/test/image/upload/v1/reference.webp",
    );
    const originalUrl = `data:application/octet-stream;base64,${"A".repeat(8 * 1024 * 1024)}`;
    const source = [{
        role: "user",
        content: [
            { type: "text", text: "反推这张图" },
            { type: "image_url", image_url: { url: originalUrl } },
        ],
    }];

    const prepared = await compressMessages(source);
    const preparedUrl = prepared[0].content[1].image_url.url;
    assert.equal(source[0].content[1].image_url.url, originalUrl);
    assert.match(preparedUrl, /^data:image\/webp;base64,/);
    assert.ok(preparedUrl.length < originalUrl.length / 3);
    assert.ok(new Blob([JSON.stringify({ messages: prepared })]).size < 3 * 1024 * 1024);
    assert.ok(encodeCalls.every(({ width, height }) => Math.max(width, height) <= 2048));
    assert.equal(disposed, true);

    const hosted = await compressMessages(source, { imgbbApiKey: "test-imgbb-key" });
    assert.equal(hosted[0].content[1].image_url.url, "https://i.ibb.co/example/text-reference.webp");
    assert.equal(source[0].content[1].image_url.url, originalUrl);
    assert.equal(uploads.length, 1);
    assert.equal(uploads[0].url, "https://api.imgbb.com/1/upload");
    assert.ok(uploads[0].form.get("image").size <= 2 * 1024 * 1024);
    assert.ok(new Blob([JSON.stringify({ messages: hosted })]).size < 1024);

    fetchedBlob = new Blob([new Uint8Array(512 * 1024)], { type: "image/webp" });
    const tudouHosted = await compressMessages(source, {
        provider: "tudou",
        referenceImageHost: "imgbb",
        imgbbApiKey: "test-imgbb-key",
    });
    assert.equal(tudouHosted[0].content[1].image_url.url, "https://res.cloudinary.com/test/image/upload/v1/reference.webp");
    assert.equal(uploads.length, 1, "Tudou must not use ImgBB, including when old settings select it");
});

test("oversized Tudou text requests no longer tell web users to switch apps", () => {
    assert.match(bundle, /assertTudouWebRequestSize=\(e,t\)=>/);
    assert.match(bundle, /ImgBB API Key/);
});

test("Tudou proxy forwards Chat Completions unchanged", async () => {
    const originalFetch = globalThis.fetch;
    const forwarded = [];
    globalThis.fetch = async (url, init) => {
        forwarded.push({ url: String(url), init });
        return new Response(JSON.stringify({
            choices: [{ message: { content: "土豆返回" } }],
        }), { status: 200, headers: { "Content-Type": "application/json" } });
    };

    try {
        const request = new Request(
            "https://www.vinsen.top/api/tudou-proxy?path=v1/chat/completions",
            {
                method: "POST",
                headers: {
                    Authorization: "Bearer test-key",
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    model: "gpt-5.5",
                    messages: [{ role: "user", content: "测试" }],
                    stream: false,
                }),
            },
        );
        const response = await tudouProxy.fetch(request);
        assert.equal(response.status, 200);
        assert.equal(forwarded.length, 1);
        assert.equal(forwarded[0].url, "https://api.ai-tudou.net/v1/chat/completions");
        const body = JSON.parse(new TextDecoder().decode(forwarded[0].init.body));
        assert.deepEqual(body.messages, [{ role: "user", content: "测试" }]);
        assert.equal(body.stream, false);
        assert.equal(await response.json().then((value) => value.choices[0].message.content), "土豆返回");
    } finally {
        globalThis.fetch = originalFetch;
    }
});

test("text requests use the global text route independently of the active image site", () => {
    assert.match(
        bundle,
        /model:n==="image"\|\|n==="text"\?ODe\(e,n,t\?\.metadata\?\.model\|\|r\)/,
    );
    assert.match(
        bundle,
        /function ODe\(e,t,n\)\{if\(t==="text"\)return textRequestConfig\(e\)\.textModel\|\|"";const r=CS\(e,t\)\|\|\[\],o=yx\(n,e\.channels,e\.activeSiteId\)/,
    );
    assert.doesNotMatch(
        bundle,
        /function ODe\(e,t,n\)\{const r=CS\(e,t\)\|\|\[\],o=yx\(t==="text"\?/,
    );
});

test("recent canvas copies and external image pastes use source priority", () => {
    const start = bundle.indexOf("K.clipboardData?.files");
    const end = bundle.indexOf('window.addEventListener("paste",O)', start);
    assert.ok(start >= 0 && end > start, "native paste handler should exist");

    const handler = bundle.slice(start, end);
    assert.match(handler, /K\.clipboardData\?\.items/);
    assert.match(handler, /f\.current\.preferInternal/);
    assert.match(handler, /Date\.now\(\)-\(f\.current\.copiedAt\|\|0\)<3e4/);
    assert.match(handler, /Ea\(Ee,lastPointerWorldRef\.current\|\|Ji\(\)\)/);
    assert.match(handler, /if\(gt\)\{K\.preventDefault\(\),Cf\(\);return\}/);
    assert.match(bundle, /if\(le&&!K\.altKey&&Q==="v"\)\{return\}/);
    assert.match(bundle, /if\(le&&!K\.altKey&&Q==="c"\)\{Ih\(\)&&\(K\.preventDefault\(\),K\.stopPropagation\(\)\);return\}/);
    assert.doesNotMatch(bundle, /returnif/);
});

test("canvas copy reads the current selection and synchronizes click selection immediately", () => {
    assert.match(bundle, /Ih=c\.useCallback\(\(\)=>\{const O=oo\.current;/);
    assert.match(bundle, /lastPointerWorldRef=c\.useRef\(null\)/);
    assert.match(bundle, /lastPointerWorldRef\.current=or\(O\.clientX,O\.clientY\)/);
    assert.match(bundle, /const K=lastPointerWorldRef\.current\|\|Ji\(\)/);
    assert.match(bundle, /copiedAt:Date\.now\(\),preferInternal:!0/);
    assert.match(bundle, /document\.addEventListener\("visibilitychange",Q\)/);
    assert.match(bundle, /ye\(gt\),oo\.current=gt;/);
    assert.match(bundle, /function sanitizeCanvasNodeClone\(e\)\{const t=\{\.\.\.e,position:\{\.\.\.e\.position\},metadata:e\.metadata\?\{\.\.\.e\.metadata\}:void 0\}/);
});

test("canvas prompt edits update the immediate copy snapshot before generation", () => {
    assert.match(
        bundle,
        /\$f=c\.useCallback\(\(O,K\)=>\{const le=vn\.current\.map\(Ee=>Ee\.id===O\?\{\.\.\.Ee,metadata:\{\.\.\.Ee\.metadata,prompt:K\}\}:Ee\);vn\.current=le,_\(Q=>Q\.map\(Ee=>Ee\.id===O\?\{\.\.\.Ee,metadata:\{\.\.\.Ee\.metadata,prompt:K\}\}:Ee\)\)\},\[\]\)/,
    );
});

test("canvas content edits update the immediate copy snapshot before generation", () => {
    assert.match(
        bundle,
        /Sv=c\.useCallback\(\(O,K\)=>\{const le=vn\.current\.map\(Ee=>Ee\.id===O\?\{\.\.\.Ee,metadata:\{\.\.\.Ee\.metadata,content:K\}\}:Ee\);vn\.current=le,_\(Q=>Q\.map\(Ee=>Ee\.id===O\?\{\.\.\.Ee,metadata:\{\.\.\.Ee\.metadata,content:K\}\}:Ee\)\)\},\[\]\)/,
    );
});

test("Ctrl/Cmd+C inside an unselected canvas prompt copies the node, not browser text", () => {
    assert.match(
        bundle,
        /if\(le&&!K\.altKey&&Q==="c"&&editable\)\{const targetText=K\.target instanceof HTMLInputElement\|\|K\.target instanceof HTMLTextAreaElement\?K\.target:null,hasSelection=!!targetText&&typeof targetText\.selectionStart==="number"&&targetText\.selectionStart!==targetText\.selectionEnd,nodeElement=Ee\?\.closest\("\[data-node-id\]"\),nodeId=nodeElement\?\.getAttribute\("data-node-id"\);if\(!hasSelection&&nodeId\)\{oo\.current=new Set\(\[nodeId\]\);Ih\(\)&&\(K\.preventDefault\(\),K\.stopPropagation\(\)\)\}return\}/,
    );
});

test("editing a generated image prompt updates the copyable node metadata", () => {
    assert.match(
        bundle,
        /const N=z=>\{\$\(z\),\(v==="image"\|\|!C\)&&n\(e\.id,z\)\}/,
    );
});

test("Alt-click duplication suppresses the browser's native save action", () => {
    assert.match(bundle, /O\.altKey&&O\.preventDefault\(\)/);
    assert.match(bundle, /onClick:_e=>\{_e\.altKey&&\(_e\.preventDefault\(\),_e\.stopPropagation\(\)\)\}/);
});

test("Alt-click duplication refreshes the synchronous canvas copy snapshot", () => {
    const start = bundle.indexOf("Rn.length&&");
    const end = bundle.indexOf("oo.current=gt", start);
    assert.ok(start >= 0 && end > start);
    assert.match(bundle.slice(start, end), /Mt=\[\.\.\.le,\.\.\.Rn\]/);
    assert.match(bundle.slice(start, end), /vn\.current=Mt,ye\(gt\)/);
});

test("pasted canvas nodes refresh the synchronous copy snapshot", () => {
    const start = bundle.indexOf("Cf=c.useCallback");
    const end = bundle.indexOf("c.useCallback(()=>{ue", start);
    assert.ok(start >= 0 && end > start);
    assert.match(bundle.slice(start, end), /return _\(tt=>\{const le=\[\.\.\.tt,\.\.\.Mt\];return vn\.current=le,le\}\)/);
});
