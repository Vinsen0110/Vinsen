import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
    inlineTudouGeminiReferences,
    isImgBbImageUrl,
    isTudouGeminiImageTarget,
} from "../tudou-reference-images.js";
import tudouProxy from "../api/tudou-proxy.js";

const target = new URL(
    "https://api.ai-tudou.net/v1beta/models/gemini-3-pro-image-preview:streamGenerateContent?alt=sse",
);
const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

test("explicit ImgBB fallback retains its temporary 2048px WebP references and output resolution", () => {
    assert.match(bundle, /TUDOU_REFERENCE_MAX_EDGE=2048/);
    assert.match(bundle, /TUDOU_REFERENCE_INITIAL_QUALITY=\.85/);
    assert.match(bundle, /TUDOU_REFERENCE_TARGET_BYTES=10\*1024\*1024/);
    assert.match(bundle, /TUDOU_REFERENCE_MAX_BYTES=14\*1024\*1024/);
    assert.match(bundle, /Math\.min\(1,o\/Math\.max\(n\.width,n\.height\)\)/, "references must not be upscaled");
    assert.match(bundle, /const l=isTudouSite\(e\),d=l\?await prepareTudouReferenceBlob\(i,n\):i/);
    assert.match(bundle, /f\.set\("image",d,l\?"reference\.webp":t\.name\|\|"reference\.png"\)/);
    assert.match(bundle, /imageSize:tudouResolution\(e\)\.toUpperCase\(\)/);
    assert.doesNotMatch(
        bundle,
        /prepareTudouReferenceBlob\([^)]*\)[^;]*metadata|metadata[^;]*prepareTudouReferenceBlob/,
        "the temporary request copy must not replace canvas metadata",
    );
});

test("recognizes Tudou Gemini image endpoints and ImgBB URLs", () => {
    assert.equal(isTudouGeminiImageTarget(target), true);
    assert.equal(isImgBbImageUrl("https://i.ibb.co/example/reference.jpg"), true);
    assert.equal(isImgBbImageUrl("https://example.com/reference.jpg"), false);
});

test("converts ImgBB fileData references to inlineData before forwarding", async () => {
    const payload = {
        contents: [{
            role: "user",
            parts: [
                { fileData: { fileUri: "https://i.ibb.co/example/reference.jpg", mimeType: "image/png" } },
                { text: "keep the same composition" },
            ],
        }],
    };
    const fetchImpl = async (url) => {
        assert.equal(url, "https://i.ibb.co/example/reference.jpg");
        return new Response(Uint8Array.from([1, 2, 3, 4]), {
            headers: {
                "Content-Length": "4",
                "Content-Type": "image/jpeg",
            },
        });
    };

    const result = await inlineTudouGeminiReferences(target, payload, { fetchImpl });

    assert.equal(result.converted, 1);
    assert.deepEqual(result.payload.contents[0].parts[0], {
        inlineData: { mimeType: "image/jpeg", data: "AQIDBA==" },
    });
    assert.equal(result.payload.contents[0].parts[1].text, "keep the same composition");
});

test("does not fetch or rewrite non-ImgBB fileData references", async () => {
    const payload = {
        contents: [{ parts: [{ fileData: { fileUri: "https://example.com/reference.jpg" } }] }],
    };
    const result = await inlineTudouGeminiReferences(target, payload, {
        fetchImpl: async () => assert.fail("unexpected fetch"),
    });

    assert.equal(result.converted, 0);
    assert.equal(result.payload.contents[0].parts[0].fileData.fileUri, "https://example.com/reference.jpg");
});

test("rejects ImgBB references above the 14 MB safety limit", async () => {
    const payload = {
        contents: [{ parts: [{ fileData: { fileUri: "https://i.ibb.co/example/too-large.webp" } }] }],
    };
    await assert.rejects(
        inlineTudouGeminiReferences(target, payload, {
            fetchImpl: async () => new Response(Uint8Array.from([1]), {
                headers: {
                    "Content-Length": String(14 * 1024 * 1024 + 1),
                    "Content-Type": "image/webp",
                },
            }),
        }),
        /exceeds the 14 MB limit/,
    );
});

test("Vercel proxy inlines one reference and submits exactly one Tudou request", async () => {
    const originalFetch = globalThis.fetch;
    let tudouRequests = 0;
    globalThis.fetch = async (url, init) => {
        const value = String(url);
        if (value.startsWith("https://i.ibb.co/")) {
            return new Response(Uint8Array.from([5, 6, 7]), {
                headers: { "Content-Type": "image/png" },
            });
        }

        tudouRequests += 1;
        assert.match(value, /api\.ai-tudou\.net\/v1beta\/models\/gemini-3-pro-image-preview:streamGenerateContent/);
        const forwarded = JSON.parse(String(init.body));
        assert.deepEqual(forwarded.contents[0].parts[0], {
            inlineData: { mimeType: "image/png", data: "BQYH" },
        });
        return new Response("data: {\"candidates\":[]}\n\n", {
            headers: { "Content-Type": "text/event-stream" },
        });
    };

    try {
        const request = new Request(
            "https://www.vinsen.top/api/tudou-proxy?path=v1beta/models/gemini-3-pro-image-preview:streamGenerateContent&alt=sse",
            {
                method: "POST",
                headers: {
                    Authorization: "Bearer test-key",
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    contents: [{
                        role: "user",
                        parts: [
                            { fileData: { fileUri: "https://i.ibb.co/example/reference.png" } },
                            { text: "keep the cat and replace only its accessories" },
                        ],
                    }],
                    generationConfig: {
                        responseModalities: ["TEXT", "IMAGE"],
                        imageConfig: { aspectRatio: "3:2", imageSize: "4K" },
                    },
                }),
            },
        );

        const response = await tudouProxy.fetch(request);
        assert.equal(response.status, 200);
        assert.equal(response.headers.get("content-type"), "text/event-stream; charset=utf-8");
        assert.equal(response.headers.get("x-tudou-references-inlined"), "pending");
        const responseBody = await response.text();
        assert.match(responseBody, /^: tudou-proxy\n\n/);
        assert.equal(tudouRequests, 1);
    } finally {
        globalThis.fetch = originalFetch;
    }
});

test("Vercel proxy sends a heartbeat before a delayed ImgBB download completes", async () => {
    const originalFetch = globalThis.fetch;
    let releaseDownload;
    let markDownloadStarted;
    const downloadStarted = new Promise((resolve) => {
        markDownloadStarted = resolve;
    });
    const downloadBlocked = new Promise((resolve) => {
        releaseDownload = resolve;
    });
    let tudouRequests = 0;

    globalThis.fetch = async (url) => {
        if (String(url).startsWith("https://i.ibb.co/")) {
            markDownloadStarted();
            await downloadBlocked;
            return new Response(Uint8Array.from([8, 9]), {
                headers: { "Content-Type": "image/webp" },
            });
        }
        tudouRequests += 1;
        return new Response("data: {\"candidates\":[]}\n\n", {
            headers: { "Content-Type": "text/event-stream" },
        });
    };

    try {
        const request = new Request(
            "https://www.vinsen.top/api/tudou-proxy?path=v1beta/models/gemini-3-pro-image-preview:streamGenerateContent&alt=sse",
            {
                method: "POST",
                headers: {
                    Authorization: "Bearer test-key",
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    contents: [{ parts: [{ fileData: { fileUri: "https://i.ibb.co/example/reference.webp" } }] }],
                    generationConfig: {
                        responseModalities: ["TEXT", "IMAGE"],
                        imageConfig: { aspectRatio: "1:1", imageSize: "4K" },
                    },
                }),
            },
        );

        const response = await tudouProxy.fetch(request);
        const reader = response.body.getReader();
        const firstChunk = await reader.read();
        assert.equal(new TextDecoder().decode(firstChunk.value), ": tudou-proxy\n\n");
        await downloadStarted;
        assert.equal(tudouRequests, 0, "Tudou must wait until the reference is inlined");

        releaseDownload();
        while (!(await reader.read()).done) {
            // Drain the upstream response so the proxy finishes cleanly.
        }
        assert.equal(tudouRequests, 1);
    } finally {
        releaseDownload?.();
        globalThis.fetch = originalFetch;
    }
});
