import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const server = await readFile(new URL("../local-preview-server.mjs", import.meta.url), "utf8");
const start = server.indexOf("async function proxyTudouGeminiStream(");
const end = server.indexOf("async function proxyTudouApi(", start);
assert.ok(start >= 0 && end > start);

async function runProxy({ chunks, json, status = 200, failRead = false }) {
    const intervals = new Map();
    const writes = [];
    let requests = 0;
    let headers;
    let index = 0;
    const tick = () => { for (const callback of intervals.values()) callback(); };
    const response = {
        writableEnded: false,
        writeHead(code, value) { headers = { code, ...value }; },
        write(value) { writes.push(Buffer.from(value)); },
        end() { this.writableEnded = true; },
    };
    const scope = {
        Buffer,
        setInterval(callback) { intervals.set(1, callback); return 1; },
        clearInterval(id) { intervals.delete(id); },
        console: { error() {} },
        errorDetails: error => error.message,
        sseData: value => `data: ${JSON.stringify(value)}\n\n`,
        inlineTudouGeminiReferences: async (target, payload) => {
            tick();
            return { payload };
        },
        fetch: async () => {
            requests++;
            tick();
            return {
                ok: status >= 200 && status < 300,
                status,
                headers: new Headers({ "content-type": json ? "application/json" : "text/event-stream" }),
                text: async () => json || "",
                body: chunks && {
                    getReader: () => ({
                        async read() {
                            // Simulate a heartbeat firing while an image frame spans network reads.
                            tick();
                            if (failRead) throw new Error("upstream disconnected");
                            return index < chunks.length
                                ? { done: false, value: Buffer.from(chunks[index++]) }
                                : { done: true };
                        },
                    }),
                },
            };
        },
    };
    const proxy = vm.runInNewContext(`${server.slice(start, end)};proxyTudouGeminiStream`, scope);
    await proxy({}, response, new URL("https://api.ai-tudou.net/v1beta/models/test:streamGenerateContent"),
        new Headers(), { contents: [] }, new AbortController());
    return { text: Buffer.concat(writes).toString(), requests, headers, intervals, response };
}

test("local proxy never injects a heartbeat inside fragmented image JSON", async () => {
    const event = `data: ${JSON.stringify({
        candidates: [{ content: { parts: [{ inlineData: { mimeType: "image/png", data: "A".repeat(8192) } }] } }],
    })}\n\n`;
    const result = await runProxy({ chunks: [event.slice(0, 300), event.slice(300, 2700), event.slice(2700)] });
    assert.equal(result.text.slice(result.text.indexOf("data:")), event);
    assert.equal(result.requests, 1);
    assert.equal(result.intervals.size, 0);
    assert.equal(result.response.writableEnded, true);
});

test("local proxy wraps a non-streaming JSON image response as one SSE event", async () => {
    const json = JSON.stringify({ candidates: [{ content: { parts: [{ inlineData: { data: "AQID" } }] } }] });
    const result = await runProxy({ json });
    assert.equal(result.text.slice(result.text.indexOf("data:")), `data: ${json}\n\n`);
    assert.equal(result.headers["Content-Type"], "text/event-stream; charset=utf-8");
    assert.equal(result.requests, 1);
});

test("local proxy exposes upstream failures and cleans up heartbeat timers", async () => {
    for (const options of [
        { json: '{"error":{"message":"provider refused"}}', status: 429 },
        { chunks: ["partial"], failRead: true },
        {},
    ]) {
        const result = await runProxy(options);
        assert.match(result.text, /data: \{"error":\{"message":/);
        assert.equal(result.intervals.size, 0);
        assert.equal(result.response.writableEnded, true);
        assert.equal(result.requests, 1);
    }
});
