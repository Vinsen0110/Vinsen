import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const vercelConfig = JSON.parse(await readFile(new URL("../vercel.json", import.meta.url), "utf8"));

function bundleBlock(start, end) {
    const startIndex = bundle.indexOf(start);
    const endIndex = bundle.indexOf(end, startIndex + start.length);
    assert.ok(startIndex >= 0 && endIndex > startIndex, `missing bundle block: ${start}`);
    return bundle.slice(startIndex, endIndex);
}

const extracted = [
    bundleBlock("function oW", "async function aW"),
    bundleBlock("function appendTudouGeminiSseEvent", "async function tudouGeminiStreamImages"),
    bundleBlock("async function tudouGeminiStreamImages", "async function qMe"),
    bundleBlock("function geminiResponseLayers", "function geminiNoImageMessage"),
    bundleBlock("function geminiNoImageMessage", "function WMe"),
    bundleBlock("function WMe", "function imageTaskIdentifier"),
].join("");

const makeRuntime = new Function("fetch", "Jq", "eW", "aW", "cn", "Tm", `${extracted};return tudouGeminiStreamImages;`);

function sseResponse(blocks, headers = {}) {
    return new Response(blocks.join("\n\n") + "\n\n", {
        headers: {
            "Content-Type": "text/event-stream",
            "X-Oneapi-Request-Id": "req-stream-test",
            ...headers,
        },
    });
}

async function runBundleStream(response) {
    let requests = 0;
    const stream = makeRuntime(
        async () => {
            requests += 1;
            return response;
        },
        () => "https://example.test/streamGenerateContent",
        () => ({ Authorization: "Bearer test-key" }),
        async () => "request failed",
        () => "image-id",
        (value) => Boolean(value) && typeof value === "object" && !Array.isArray(value),
    );
    const result = await stream({ model: "test" }, { contents: [] });
    return { requests, result };
}

test("Tudou empty streams retain upstream diagnostics without replaying the POST", () => {
    assert.match(bundle, /function tudouStreamRequestId\(e\)/);
    assert.match(bundle, /provider:"tudou",requestId:o/);
    assert.match(bundle, /eventCount:d\.length/);
    assert.match(bundle, /Tudou \\u6D41\\u5F0F\\u54CD\\u5E94\\u683C\\u5F0F\\u5F02\\u5E38/);
    assert.match(bundle, /request-id=/);
    assert.match(bundle, /finishReason=/);
    assert.match(bundle, /async function dW\(e,t,n,r,o\)\{o\?\.onProgress\?\.\(\{progress:10,stage:"generating"\}\)/);

    const start = bundle.indexOf("async function tudouGeminiStreamImages");
    const streamFunction = bundle.slice(start, bundle.indexOf("async function qMe", start));
    assert.equal((streamFunction.match(/fetch\(/g) || []).length, 1);
});

test("actual bundle combines later image chunks and sends one POST", async () => {
    const { requests, result } = await runBundleStream(sseResponse([
        'data: {"candidates":[{"content":{"parts":[{"text":"working"}]}}]}',
        'data: {"candidates":[{"content":{"parts":[{"inline_data":{"mime_type":"image/png","data":"AQID"}}]},"finishReason":"STOP"}]}',
        "data: [DONE]",
    ]));
    assert.equal(requests, 1);
    assert.deepEqual(result, [{ id: "image-id", dataUrl: "data:image/png;base64,AQID" }]);
});

test("actual bundle accepts adjacent complete data lines without blank separators", async () => {
    const { requests, result } = await runBundleStream(new Response(
        ': keepalive\r\ndata: {"candidates":[{"content":{"parts":[{"text":"working"}]}}]}\r\n'
        + 'data: {"candidates":[{"content":{"parts":[{"inlineData":{"mimeType":"image/png","data":"AQID"}}]}}]}\r\n'
        + 'data: [DONE]\r\n',
        { headers: { "Content-Type": "text/event-stream" } },
    ));
    assert.equal(requests, 1);
    assert.equal(result[0].dataUrl, "data:image/png;base64,AQID");
});

test("actual bundle retains standard multiline SSE JSON and comments", async () => {
    const { result } = await runBundleStream(sseResponse([
        ': heartbeat\ndata: {\ndata: "candidates":[{"content":{"parts":[{"inlineData":{"data":"AQID","mimeType":"image/png"}}]}}]\ndata: }',
    ]));
    assert.equal(result[0].dataUrl, "data:image/png;base64,AQID");
});

test("actual bundle assembles fragmented UTF-8 and large inline images without replay", async () => {
    const data = "AQID".repeat(100000);
    const raw = new TextEncoder().encode(
        `data: ${JSON.stringify({ candidates: [{ content: { parts: [{ text: "测试" }, { inlineData: { data, mimeType: "image/png" } }] } }] })}\n\n`,
    );
    const response = new Response(new ReadableStream({
        start(controller) {
            for (let offset = 0; offset < raw.length; offset += 997) controller.enqueue(raw.slice(offset, offset + 997));
            controller.close();
        },
    }), { headers: { "Content-Type": "text/event-stream" } });
    const { requests, result } = await runBundleStream(response);
    assert.equal(requests, 1);
    assert.equal(result[0].dataUrl, `data:image/png;base64,${data}`);
});

test("malformed JSON is not silently discarded beside a valid event", async () => {
    await assert.rejects(runBundleStream(sseResponse([
        'data: {broken}\ndata: {"candidates":[{"content":{"parts":[{"inlineData":{"data":"AQID"}}]}}]}',
    ])), /Tudou 流式响应格式异常/);
});

test("actual bundle explains an N/A stream with upstream diagnostics", async () => {
    await assert.rejects(
        runBundleStream(sseResponse([
            'data: {"candidates":[{"content":{"parts":[{"text":"No image was generated"}]},"finishReason":"OTHER","finishMessage":"upstream returned no image"}]}',
            "data: [DONE]",
        ])),
        (error) => {
            assert.match(error.message, /^Tudou Gemini 流已结束，但上游没有产出图片/);
            assert.match(error.message, /finishReason=OTHER/);
            assert.match(error.message, /finishMessage=upstream returned no image/);
            assert.match(error.message, /text=No image was generated/);
            assert.match(error.message, /request-id=req-stream-test/);
            assert.match(error.message, /events=1/);
            return true;
        },
    );
});

test("actual bundle surfaces errors that arrive after HTTP 200", async () => {
    await assert.rejects(
        runBundleStream(sseResponse(['data: {"error":{"message":"provider timeout"}}'])),
        /provider timeout/,
    );
    await assert.rejects(
        runBundleStream(sseResponse(["data: {broken}"])),
        /Tudou 流式响应格式异常/,
    );
});

test("model SVG icons are served from browser cache after their first load", () => {
    const rule = vercelConfig.headers.find((entry) => entry.source === "/icons/:path*");
    assert.ok(rule);
    assert.deepEqual(rule.headers, [{
        key: "Cache-Control",
        value: "public, max-age=31536000, immutable",
    }]);
});
