import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
    APILIO_FILE_MAX_BYTES,
    APILIO_REFERENCE_TARGET_BYTES,
    APILIO_UPLOAD_TIMEOUT_MS,
    apilioHostedReferenceUrl,
    apilioFilesEndpoint,
    isApilioHostedImageUrl,
    prepareApilioReferenceBlob,
    uploadApilioReferenceBlob,
} from "../apilio-reference-upload.js";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

test("builds the documented Apilio files endpoint", () => {
    assert.equal(apilioFilesEndpoint("https://api.apilio.ai"), "https://api.apilio.ai/v1/files");
    assert.equal(apilioFilesEndpoint("https://example.com/v1/"), "https://example.com/v1/files");
    assert.equal(APILIO_FILE_MAX_BYTES, 20 * 1024 * 1024);
    assert.equal(APILIO_REFERENCE_TARGET_BYTES, 18 * 1024 * 1024);
    assert.equal(APILIO_UPLOAD_TIMEOUT_MS, 3 * 60 * 1000);
});

test("recognizes only trusted HTTPS Apilio image hosts", () => {
    assert.equal(isApilioHostedImageUrl("https://webstatic.aiproxy.vip/output/image.png"), true);
    assert.equal(isApilioHostedImageUrl("https://files.closeai.fans/filesystem/uploads/image.png"), true);
    assert.equal(isApilioHostedImageUrl("https://cdn.gptbest.vip/file/image.png"), true);
    assert.equal(isApilioHostedImageUrl("https://imgbb.com/image.png"), false);
    assert.equal(isApilioHostedImageUrl("https://cdn.gptbest.vip.evil.example/image.png"), false);
    assert.equal(isApilioHostedImageUrl("http://cdn.gptbest.vip/file/image.png"), false);
});

test("prefers the preserved Apilio remote source over a local preview", () => {
    assert.equal(apilioHostedReferenceUrl({
        dataUrl: "blob:local-preview",
        remoteSourceUrl: "https://webstatic.aiproxy.vip/output/generated.png",
    }), "https://webstatic.aiproxy.vip/output/generated.png");
    assert.equal(apilioHostedReferenceUrl({
        dataUrl: "blob:local-preview",
        url: "https://cdn.gptbest.vip/file/uploaded.png",
    }), "https://cdn.gptbest.vip/file/uploaded.png");
    assert.equal(apilioHostedReferenceUrl({ dataUrl: "https://i.ibb.co/image.png" }), "");
});

test("uploads a small original Blob with only the documented multipart file field", async () => {
    const original = new Blob(["original-image"], { type: "image/png" });
    const progress = [];
    let request;
    const url = await uploadApilioReferenceBlob(
        { baseUrl: "https://api.apilio.ai/v1", apiKey: "test-key" },
        original,
        {
            filename: "一双拖鞋.png",
            onProgress: (value) => progress.push(value),
            fetchImpl: async (target, init) => {
                request = { target, init };
                return new Response(JSON.stringify({
                    id: "file-test",
                    object: "file",
                    bytes: original.size,
                    filename: "original.png",
                    url: "https://cdn.gptbest.vip/file/reference.png",
                }), { status: 200 });
            },
        },
    );

    assert.equal(url, "https://cdn.gptbest.vip/file/reference.png");
    assert.equal(request.target, "https://api.apilio.ai/v1/files");
    assert.equal(request.init.method, "POST");
    assert.deepEqual(request.init.headers, { Authorization: "Bearer test-key" });
    assert.deepEqual([...request.init.body.keys()], ["file"]);
    assert.equal(request.init.body.get("file").size, original.size);
    assert.equal(request.init.body.get("file").name, "reference.png");
    assert.deepEqual(progress, [
        { progress: 2, stage: "uploading" },
        { progress: 10, stage: "uploading" },
    ]);
});

test("compresses an oversized request copy at original dimensions first", async () => {
    const original = new Blob([new Uint8Array(21)]);
    const reference = { url: "blob:preview", storageKey: "image:original", width: 4096, height: 3072 };
    const snapshot = structuredClone(reference);
    const calls = [];
    const compressed = await prepareApilioReferenceBlob(original, {
        targetBytes: 10,
        maxBytes: 20,
        decodeImage: async () => ({ image: {}, width: 4096, height: 3072 }),
        encodeImage: async (_image, width, height, quality) => {
            calls.push({ width, height, quality });
            return new Blob([new Uint8Array(9)], { type: "image/webp" });
        },
    });

    assert.notEqual(compressed, original);
    assert.equal(compressed.size, 9);
    assert.deepEqual(calls[0], { width: 4096, height: 3072, quality: 0.92 });
    assert.deepEqual(reference, snapshot, "temporary compression must not rewrite preview metadata");
});

test("keeps files at or below the upload limit unchanged", async () => {
    const original = new Blob([new Uint8Array(20)], { type: "image/png" });
    let decoded = false;
    const prepared = await prepareApilioReferenceBlob(original, {
        targetBytes: 18,
        maxBytes: 20,
        decodeImage: async () => {
            decoded = true;
            throw new Error("must not decode");
        },
    });

    assert.equal(prepared, original);
    assert.equal(decoded, false);
});

test("reduces dimensions only after all full-size quality attempts remain too large", async () => {
    const original = new Blob([new Uint8Array(21)]);
    const calls = [];
    const compressed = await prepareApilioReferenceBlob(original, {
        targetBytes: 10,
        maxBytes: 20,
        decodeImage: async () => ({ image: {}, width: 4000, height: 3000 }),
        encodeImage: async (_image, width, height, quality) => {
            calls.push({ width, height, quality });
            return new Blob([new Uint8Array(width === 4000 ? 12 : 8)], { type: "image/webp" });
        },
    });

    assert.equal(compressed.size, 8);
    assert.equal(calls.slice(0, 8).every(({ width, height }) => width === 4000 && height === 3000), true);
    assert.deepEqual(calls[8], { width: 3600, height: 2700, quality: 0.92 });
});

test("rejects an upload response without the documented top-level url", async () => {
    await assert.rejects(
        uploadApilioReferenceBlob(
            { baseUrl: "https://api.apilio.ai", apiKey: "test-key" },
            new Blob(["image"], { type: "image/png" }),
            { fetchImpl: async () => new Response(JSON.stringify({ id: "file-test" }), { status: 200 }) },
        ),
        /没有返回有效 URL/,
    );
});

test("fails a stalled Apilio upload instead of waiting forever", async () => {
    await assert.rejects(
        uploadApilioReferenceBlob(
            { baseUrl: "https://api.apilio.ai", apiKey: "test-key" },
            new Blob(["image"], { type: "image/png" }),
            {
                timeoutMs: 5,
                fetchImpl: async (_target, { signal }) => new Promise((_resolve, reject) => {
                    signal.addEventListener("abort", () => reject(new DOMException("Aborted", "AbortError")), { once: true });
                }),
            },
        ),
        /上传超时/,
    );
});

test("all Apilio reference branches use the Apilio file uploader", () => {
    assert.match(bundle, /apilioHostedReferenceUrl.*uploadApilioReferenceBlob.*from"\.\.\/apilio-reference-upload\.js"/);
    assert.doesNotMatch(bundle, /if\(isApilioSite\(e\)\)\{const [a-z]=apilioHostedReferenceUrl\(t\);if\([a-z]\)return reportImageTaskProgress\(r,10\),[a-z]\}/);
    assert.match(bundle, /if\(isApilioSite\(e\)\)return uploadApilioReferenceBlob\(e,i,\{filename:t\.name\|\|"reference\.png",signal:n,onProgress:r\}\)/);
    assert.match(bundle, /if\(isApilioSite\(t\)\)return uploadApilioReferenceBlob\(t,l,\{filename:"text-reference",signal:n\}\)/);
    assert.match(bundle, /if\(isRunningHubSite\(t\)\)return uploadRunningHubReferenceBlob\(t,l,\{signal:n\}\)/);
    assert.match(bundle, /reportImageTaskProgress\(i,10,"generating"\)/);
    assert.match(bundle, /generationStage:[a-z]/);
    assert.match(bundle, /remoteSourceUrl:isRemoteImageUrl\(e\)\?e:void 0/);
    assert.match(bundle, /remoteSourceUrl:e\.metadata\.remoteSourceUrl/);
    assert.match(bundle, /async function q6\(e,t\).*dataUrl:await n\(r\)/);
    assert.doesNotMatch(bundle, /async function q6\(e,t\).*apilioHostedReferenceUrl\(r\)/);
    assert.match(bundle, /K==="image"&&isApilioSite\(Se\)/);
    assert.match(bundle, /O\.type===Ne\.Image&&isApilioSite\(Se\)/);
    assert.match(bundle, /submitApilioNanoImages\(a,l,d,i,o\?\.signal,o\?\.onProgress\)/);
    assert.match(bundle, /submitImageRequest\(e,"\/images\/generations",i,wA\(e,"application\/json"\),o,!0,a\)/);
    assert.match(bundle, /function apilioNanoImagePayload\(e,t,n=\[\]\).*model:"nano-banana-pro".*aspect_ratio:o.*image:n.*image_size:r.*prompt:Zq\(e,t\),response_format:"url"/);
    assert.match(bundle, /f\.set\("key",a\),f\.set\("image",d,l\?"reference\.webp":t\.name\|\|"reference\.png"\)/, "explicit ImgBB fallback must remain available");
});

test("localhost keeps Chrome's native directory picker for real local projects", async () => {
    const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");
    assert.match(indexHtml, /var isPreview = new URLSearchParams\(location\.search\)\.has\("preview"\);/);
    assert.match(indexHtml, /if \(\(window\.oldHouseDesktop \|\| typeof window\.showDirectoryPicker === "function"\) && !isPreview\) return;/);
    assert.doesNotMatch(indexHtml, /location\.hostname !== "127\.0\.0\.1"/);
});
