import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { createServer } from "node:net";
import { fileURLToPath } from "node:url";
import {
    CLOUDINARY_MAX_BYTES, CLOUDINARY_TARGET_BYTES, CLOUDINARY_MAX_PIXELS,
    usesCloudinaryReferenceHost, isCloudinaryImageUrl, isLegacyImgBbImageUrl,
    prepareCloudinaryReferenceBlob, uploadCloudinaryReferenceBlob, cloudinaryReferenceSource,
} from "../cloudinary-reference-upload.js";
import { createCloudinarySignatureHandler } from "../api/cloudinary-signature.js";
import { inlineTudouGeminiReferences } from "../tudou-reference-images.js";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const token = "test-upload-token-with-at-least-32-characters";
const env = {
    CLOUDINARY_CLOUD_NAME: "test-cloud", CLOUDINARY_API_KEY: "123456",
    CLOUDINARY_API_SECRET: "server-secret-never-in-client", CLOUDINARY_UPLOAD_PRESET: "test_preset",
    CLOUDINARY_UPLOAD_TOKEN_HASHES: createHash("sha256").update(token).digest("hex"),
};
const config = { provider: "tudou", cloudinaryUploadToken: token };
const input = new Blob(["original-image"], { type: "image/png" });
const decodeImage = async () => ({ image: {}, width: 4096, height: 4096, dispose() {} });
const request = (key = token, init = {}) => new Request("https://www.vinsen.top/api/cloudinary-signature", {
    method: "POST", headers: { Authorization: `Bearer ${key}` }, ...init,
});
function signer(options = {}) {
    return createCloudinarySignatureHandler({ env, rateWindows: new Map(), uuid: () => "test-id", ...options });
}
async function signedBody() {
    return (await signer()(request())).json();
}
function responseImage(signed) {
    return {
        public_id: signed.params.public_id,
        secure_url: `https://res.cloudinary.com/${signed.cloudName}/image/upload/v1/${signed.params.public_id}.png`,
    };
}

test("only Tudou and GRSAI use Cloudinary, regardless of saved legacy host preferences", () => {
    for (const provider of ["tudou", "grsai"]) {
        assert.equal(usesCloudinaryReferenceHost({ provider }), true);
        assert.equal(usesCloudinaryReferenceHost({ provider, referenceImageHost: "imgbb" }), true);
    }
    for (const provider of ["apilio", "runninghub", "apimart", "default", "custom"]) {
        assert.equal(usesCloudinaryReferenceHost({ provider }), false);
    }
    assert.equal(usesCloudinaryReferenceHost({ baseUrl: "https://api.ai-tudou.net/v1" }), true);
    assert.equal(usesCloudinaryReferenceHost({ baseUrl: "https://grsaiapi.com/v1" }), true);
    assert.equal(usesCloudinaryReferenceHost({ baseUrl: "https://evil.test/api.ai-tudou.net" }), false);
    assert.equal(usesCloudinaryReferenceHost({ provider: "apilio", baseUrl: "https://api.ai-tudou.net" }), false);
});

test("host validators reject insecure, credentialed, non-image and lookalike URLs", () => {
    assert.equal(isCloudinaryImageUrl("https://res.cloudinary.com/demo/image/upload/v1/image.png"), true);
    for (const url of [
        "http://res.cloudinary.com/demo/image/upload/a.png",
        "https://res.cloudinary.com.evil.test/demo/image/upload/a.png",
        "https://res.cloudinary.com/demo/video/upload/a.mp4",
        "https://user:pass@res.cloudinary.com/demo/image/upload/a.png",
        "https://res.cloudinary.com:8080/demo/image/upload/a.png",
        "https://res.cloudinary.com/demo/image/fetch/https://example.com/a.png",
    ]) assert.equal(isCloudinaryImageUrl(url), false, url);
    assert.equal(isLegacyImgBbImageUrl("https://i.ibb.co/a/b.png"), true);
    assert.equal(isLegacyImgBbImageUrl("https://i.ibb.co.evil.test/a.png"), false);
});

test("a compliant 4K original is returned by identity with no encoding or image mutation", async () => {
    let disposed = false;
    const before = await input.arrayBuffer();
    const result = await prepareCloudinaryReferenceBlob(input, {
        decodeImage: async () => ({ image: {}, width: 4096, height: 4096, dispose: () => { disposed = true; } }),
        encodeImage: () => assert.fail("a compliant original must not be compressed"),
    });
    assert.equal(result, input);
    assert.deepEqual(await input.arrayBuffer(), before);
    assert.equal(disposed, true);
});

test("overweight references are compressed at original dimensions first", async () => {
    const original = new Blob([new Uint8Array(CLOUDINARY_MAX_BYTES + 1)], { type: "image/png" });
    const calls = [];
    const result = await prepareCloudinaryReferenceBlob(original, {
        decodeImage,
        encodeImage: async (_image, width, height, quality) => {
            calls.push({ width, height, quality });
            return new Blob(["copy"], { type: "image/webp" });
        },
    });
    assert.notEqual(result, original);
    assert.equal(original.size, CLOUDINARY_MAX_BYTES + 1);
    assert.deepEqual(calls, [{ width: 4096, height: 4096, quality: 0.94 }]);
    assert.equal(result.type, "image/webp");
});

test("too many pixels triggers proportionate downscaling even for a small file", async () => {
    let dimensions;
    await prepareCloudinaryReferenceBlob(input, {
        decodeImage: async () => ({ image: {}, width: 9000, height: 6000 }),
        encodeImage: async (_image, width, height) => {
            dimensions = { width, height };
            return new Blob(["copy"], { type: "image/webp" });
        },
    });
    assert.ok(dimensions.width * dimensions.height <= CLOUDINARY_MAX_PIXELS);
    assert.ok(Math.abs(dimensions.width / dimensions.height - 1.5) < 0.001);
});

test("compression tries quality changes before reducing dimensions, and never chooses JPEG", async () => {
    const original = new Blob([new Uint8Array(CLOUDINARY_MAX_BYTES + 1)], { type: "image/png" });
    const calls = [];
    await prepareCloudinaryReferenceBlob(original, {
        decodeImage,
        encodeImage: async (_image, width, height, quality) => {
            calls.push({ width, height, quality });
            return { size: calls.length <= 5 ? CLOUDINARY_TARGET_BYTES + 1 : 100, type: "image/webp" };
        },
    });
    assert.equal(calls.length, 6);
    assert.ok(calls.slice(0, 5).every(item => item.width === 4096 && item.height === 4096));
    assert.ok(calls[5].width < 4096);
});

test("failed compression is bounded and does not overwrite the original", async () => {
    const original = new Blob([new Uint8Array(CLOUDINARY_MAX_BYTES + 1)], { type: "image/png" });
    let calls = 0, disposed = false;
    await assert.rejects(prepareCloudinaryReferenceBlob(original, {
        decodeImage: async () => ({ image: {}, width: 1200, height: 1200, dispose: () => { disposed = true; } }),
        encodeImage: async () => { calls++; return { size: CLOUDINARY_MAX_BYTES + 1, type: "image/webp" }; },
    }), /本地原图未修改/);
    assert.ok(calls <= 60);
    assert.equal(disposed, true);
    assert.equal(original.size, CLOUDINARY_MAX_BYTES + 1);
});

test("empty or unsupported input fails before decoding or uploading", async () => {
    for (const blob of [new Blob(), new Blob(["x"], { type: "image/svg+xml" }), new Blob(["x"], { type: "image/gif" })]) {
        await assert.rejects(prepareCloudinaryReferenceBlob(blob, {
            decodeImage: () => assert.fail("unexpected decode"),
        }));
    }
});

test("cancellation releases decoded pixels without uploading", async () => {
    const controller = new AbortController();
    let disposed = false;
    await assert.rejects(prepareCloudinaryReferenceBlob(input, {
        signal: controller.signal,
        decodeImage: async () => {
            controller.abort();
            return { image: {}, width: 20, height: 20, dispose: () => { disposed = true; } };
        },
    }), { name: "AbortError" });
    assert.equal(disposed, true);
});

test("signed upload sends original bytes and only upload parameters, never provider or application secrets", async () => {
    const signed = await signedBody();
    const calls = [], progress = [];
    const url = await uploadCloudinaryReferenceBlob({ ...config, apiKey: "provider-secret" }, input, {
        decodeImage, onProgress: value => progress.push(value),
        fetchImpl: async (url, init) => {
            calls.push({ url, init });
            if (calls.length === 1) {
                assert.equal(url, "/api/cloudinary-signature");
                assert.equal(init.headers.Authorization, `Bearer ${token}`);
                return Response.json(signed);
            }
            assert.equal(url, "https://api.cloudinary.com/v1_1/test-cloud/image/upload");
            assert.equal(init.headers, undefined);
            assert.equal(init.credentials, "omit");
            const body = init.body;
            assert.equal(body.get("api_key"), "123456");
            assert.equal(body.get("signature"), signed.signature);
            assert.equal(body.get("overwrite"), "false");
            assert.deepEqual(await body.get("file").arrayBuffer(), await input.arrayBuffer());
            assert.equal(body.get("file").name, "reference.png");
            assert.equal(body.has("api_secret"), false);
            assert.equal(body.has("Authorization"), false);
            return Response.json(responseImage(signed));
        },
    });
    assert.equal(calls.length, 2);
    assert.equal(url, responseImage(signed).secure_url);
    assert.deepEqual(progress.map(value => value.stage), ["preparing", "uploading", "uploading"]);
});

test("missing upload credential stops before reading pixels or contacting any service", async () => {
    await assert.rejects(uploadCloudinaryReferenceBlob({}, input, {
        decodeImage: () => assert.fail("unexpected decode"), fetchImpl: () => assert.fail("unexpected request"),
    }), /上传服务访问码/);
});

test("signature failures do not fall back to ImgBB or submit a model request", async () => {
    let calls = 0;
    await assert.rejects(uploadCloudinaryReferenceBlob(config, input, {
        decodeImage, fetchImpl: async () => { calls++; return Response.json({ error: { message: "denied" } }, { status: 401 }); },
    }), /denied/);
    assert.equal(calls, 1);
});

test("upload rejects expired signatures, bad destinations, and arbitrary signed fields", async () => {
    for (const mutate of [
        data => { data.cloudName = "evil/path"; },
        data => { data.params.timestamp = 0; },
        data => { data.params.overwrite = true; },
        data => { data.params.transformation = "w_1"; },
    ]) {
        let calls = 0;
        const signed = await signedBody();
        mutate(signed);
        await assert.rejects(uploadCloudinaryReferenceBlob(config, input, {
            decodeImage, fetchImpl: async () => { calls++; return Response.json(signed); },
        }), /签名/);
        assert.equal(calls, 1);
    }
});

test("upload rejects insecure or unrelated result URLs and mismatched asset identifiers", async () => {
    for (const mutate of [
        result => { result.secure_url = "https://evil.test/reference.png"; },
        result => { result.secure_url = result.secure_url.replace("test-cloud", "other-cloud"); },
        result => { result.public_id = "different-asset"; },
    ]) {
        const signed = await signedBody();
        const result = responseImage(signed);
        mutate(result);
        let count = 0;
        await assert.rejects(uploadCloudinaryReferenceBlob(config, input, {
            decodeImage, fetchImpl: async () => Response.json(++count === 1 ? signed : result),
        }), /有效的参考图链接/);
    }
});

test("stalled uploads time out without an automatic paid retry", async () => {
    const signed = await signedBody();
    let calls = 0;
    await assert.rejects(uploadCloudinaryReferenceBlob(config, input, {
        decodeImage, timeoutMs: 20,
        fetchImpl: async (_url, init) => {
            if (++calls === 1) return Response.json(signed);
            return new Promise((_resolve, reject) => init.signal.addEventListener("abort", () => reject(init.signal.reason)));
        },
    }), /上传超时/);
    assert.equal(calls, 2);
});

test("old ImgBB and Cloudinary references remain usable with no upload or metadata mutation", async () => {
    for (const url of ["https://i.ibb.co/example/reference.png", "https://res.cloudinary.com/demo/image/upload/v1/reference.png"]) {
        const reference = Object.freeze({ dataUrl: url, storageKey: "unchanged-original" });
        assert.equal(await cloudinaryReferenceSource({ provider: "tudou" }, reference, {
            readStoredBlob: () => assert.fail("unexpected local write/read"),
            fetchImpl: () => assert.fail("unexpected upload"),
        }), url);
        assert.equal(reference.storageKey, "unchanged-original");
    }
});

test("stored original is preferred and reference metadata never gets overwritten with the upload result", async () => {
    const signed = await signedBody();
    const reference = Object.freeze({ dataUrl: "blob:preview", storageKey: "original-key" });
    let calls = 0;
    await cloudinaryReferenceSource(config, reference, {
        decodeImage, readStoredBlob: async key => { assert.equal(key, "original-key"); return input; },
        fetchImpl: async () => Response.json(++calls === 1 ? signed : responseImage(signed)),
    });
    assert.equal(reference.dataUrl, "blob:preview");
    assert.equal(reference.storageKey, "original-key");
    assert.equal(calls, 2);
});

test("signer fails closed until credentials and authorized token hashes are configured", async () => {
    const handler = signer({ env: {} });
    const result = await handler(request());
    assert.equal(result.status, 503);
    assert.match((await result.json()).error.message, /未配置/);
    assert.equal((await signer()(request("wrong-token-with-at-least-32-characters"))).status, 401);
});

test("signer signs only server-owned parameters using SHA-256 and exposes no secret", async () => {
    const handler = signer({ now: () => 1_800_000_000_000 });
    const result = await handler(request(token, {
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ public_id: "victim", overwrite: true, transformation: "w_1" }),
    }));
    const data = await result.json();
    const serialized = Object.keys(data.params).sort().map(key => `${key}=${data.params[key]}`).join("&");
    assert.equal(data.signature, createHash("sha256").update(serialized + env.CLOUDINARY_API_SECRET).digest("hex"));
    assert.equal(data.params.overwrite, false);
    assert.equal(data.params.public_id, "laowu-reference/test-id");
    assert.equal(data.params.timestamp, 1_800_000_000);
    assert.equal(data.params.transformation, undefined);
    assert.equal(JSON.stringify(data).includes(env.CLOUDINARY_API_SECRET), false);
    assert.equal(JSON.stringify(data).includes(token), false);
    assert.equal(result.headers.get("cache-control"), "no-store");
});

test("signer restricts browser origins and preflight, while allowing credentialed server callers", async () => {
    const handler = signer();
    assert.equal((await handler(request(token, { method: "GET" }))).status, 405);
    assert.equal((await handler(request(token, {
        headers: { Authorization: `Bearer ${token}`, Origin: "https://evil.test" },
    }))).status, 403);
    const result = await handler(request(token, {
        method: "OPTIONS", headers: { Origin: "https://vinsen0110.github.io" },
    }));
    assert.equal(result.status, 200);
    assert.equal(result.headers.get("access-control-allow-origin"), "https://vinsen0110.github.io");
});

test("signer rate-limits authorized token bursts and resets the window", async () => {
    let time = 1_800_000_000_000;
    const handler = signer({ now: () => time });
    for (let i = 0; i < 20; i++) assert.equal((await handler(request())).status, 200);
    assert.equal((await handler(request())).status, 429);
    time += 60_001;
    assert.equal((await handler(request())).status, 200);
});

function between(start, end) {
    const first = bundle.indexOf(start);
    const last = bundle.indexOf(end, first + start.length);
    assert.ok(first >= 0 && last > first, start);
    return bundle.slice(first, last);
}

test("actual bundle image uploader routes only the two target providers to Cloudinary", async () => {
    const code = between("imgbbReferenceSource=async function(", ";\nconst TEXT_REFERENCE_MAX_EDGE");
    const calls = [];
    const upload = vm.runInNewContext(`${code};imgbbReferenceSource`, {
        usesCloudinaryReferenceHost, cloudinaryReferenceSource: async (...args) => { calls.push(["cloudinary", args[0].provider]); return "cloud"; },
        Sw: async () => input, Xl: async () => "",
        isImgBbReferenceUrl: isLegacyImgBbImageUrl,
        isTudouSite: e => e.provider === "tudou", isApilioSite: e => e.provider === "apilio",
        uploadApilioReferenceBlob: async () => { calls.push(["apilio"]); return "apilio"; },
        prepareTudouReferenceBlob: async blob => blob, TUDOU_REFERENCE_MAX_BYTES: 14 * 1024 * 1024,
        FormData, reportImageTaskProgress() {},
        Ln: { post: async () => { calls.push(["imgbb"]); return { data: { success: true, data: { url: "https://i.ibb.co/test/a.png" } } }; } },
    });
    for (const provider of ["tudou", "grsai", "apilio", "custom"]) {
        await upload({ provider, imgbbApiKey: "legacy" }, { storageKey: "key" });
    }
    await upload({ provider: "tudou", referenceImageHost: "imgbb", imgbbApiKey: "legacy" }, { storageKey: "key" });
    assert.deepEqual(calls, [["cloudinary", "tudou"], ["cloudinary", "grsai"], ["apilio"], ["imgbb"], ["cloudinary", "tudou"]]);
});

test("settings expose only the application upload access code, not ImgBB or Cloudinary secrets", () => {
    const settings = between('className:"api-settings-page"', 'className:"api-site-tabs"');
    assert.match(settings, /上传服务访问码/);
    assert.doesNotMatch(settings, /ImgBB|imgbbApiKey|referenceImageHost|Cloudinary 上传凭证|cloudinaryApiSecret/);
});

test("actual text-reference branch changes only Tudou/GRSAI while keeping native uploaders", async () => {
    const code = between("async function prepareTextReferenceDataUrl(", "async function prepareTextChatMessages(");
    const calls = [];
    const prepare = vm.runInNewContext(`${code};prepareTextReferenceDataUrl`, {
        usesCloudinaryReferenceHost,
        cloudinaryReferenceSource: async e => { calls.push(e.provider); return "cloud"; },
        isApilioSite: e => e.provider === "apilio", isRunningHubSite: e => e.provider === "runninghub",
        fetch: async () => new Response(input), throwIfTudouReferenceAborted() {},
        uploadApilioReferenceBlob: async () => "native-apilio",
        uploadRunningHubReferenceBlob: async () => "native-rh",
        prepareTextReferenceBlob: async blob => blob, P$e: async () => "inline",
        textReferenceImgBbKey: () => "", uploadTextReferenceBlob: () => assert.fail("unexpected fallback"),
    });
    assert.equal(await prepare("blob:test", { provider: "tudou" }, null, 100), "cloud");
    assert.equal(await prepare("blob:test", { provider: "grsai" }, null, 100), "cloud");
    assert.equal(await prepare("blob:test", { provider: "apilio" }, null, 100), "native-apilio");
    assert.equal(await prepare("blob:test", { provider: "runninghub" }, null, 100), "native-rh");
    assert.equal(await prepare("blob:test", { provider: "custom" }, null, 100), "inline");
    assert.deepEqual(calls, ["tudou", "grsai"]);
});

test("Tudou proxy inlines Cloudinary references as MIME-correct Base64 and preserves prompt and generation parameters", async () => {
    const payload = {
        contents: [{ parts: [{ fileData: { fileUri: "https://res.cloudinary.com/demo/image/upload/v1/a.png" } }, { text: "prompt" }] }],
        generationConfig: { imageConfig: { imageSize: "4K", aspectRatio: "16:9" } },
    };
    const result = await inlineTudouGeminiReferences(
        new URL("https://api.ai-tudou.net/v1beta/models/test:generateContent"), payload,
        { fetchImpl: async (_url, init) => {
            assert.equal(init.redirect, "manual");
            return new Response(new Uint8Array([1, 2]), { headers: { "Content-Type": "image/png" } });
        } },
    );
    assert.equal(result.converted, 1);
    assert.deepEqual(result.payload.contents[0].parts[0], { inlineData: { mimeType: "image/png", data: "AQI=" } });
    assert.equal(result.payload.contents[0].parts[1].text, "prompt");
    assert.equal(result.payload.generationConfig.imageConfig.imageSize, "4K");
});

test("Tudou proxy rejects external redirects before making another request", async () => {
    let calls = 0;
    await assert.rejects(inlineTudouGeminiReferences(
        new URL("https://api.ai-tudou.net/v1beta/models/test:generateContent"),
        { contents: [{ parts: [{ fileData: { fileUri: "https://res.cloudinary.com/demo/image/upload/v1/a.png" } }] }] },
        { fetchImpl: async () => { calls++; return new Response(null, { status: 302, headers: { Location: "http://127.0.0.1/secret" } }); } },
    ), /unsupported host/);
    assert.equal(calls, 1);
});

test("Tudou proxy bounds streaming downloads even when Content-Length is absent", async () => {
    await assert.rejects(inlineTudouGeminiReferences(
        new URL("https://api.ai-tudou.net/v1beta/models/test:generateContent"),
        { contents: [{ parts: [{ fileData: { fileUri: "https://res.cloudinary.com/demo/image/upload/v1/a.png" } }] }] },
        { fetchImpl: async () => new Response(new Uint8Array(14 * 1024 * 1024 + 1), {
            headers: { "Content-Type": "image/png" },
        }) },
    ), /14 MB limit/);
});

test("local/desktop signature route signs with server configuration and denies cross-origin requests", { timeout: 10_000 }, async () => {
    const allocator = createServer();
    allocator.listen(0, "127.0.0.1");
    await once(allocator, "listening");
    const { port } = allocator.address();
    await new Promise(resolve => allocator.close(resolve));
    const child = spawn(process.execPath, ["local-preview-server.mjs"], {
        cwd: fileURLToPath(new URL("../", import.meta.url)),
        env: { ...process.env, ...env, PORT: String(port), HOST: "127.0.0.1" },
        stdio: ["ignore", "pipe", "pipe"],
    });
    const exit = once(child, "exit");
    try {
        await Promise.race([
            once(child.stdout, "data"),
            exit.then(() => { throw new Error("Local server exited before listening"); }),
        ]);
        const origin = `http://127.0.0.1:${port}`;
        const result = await fetch(`${origin}/api/cloudinary-signature`, {
            method: "POST", headers: { Origin: origin, Authorization: `Bearer ${token}` },
        });
        assert.equal(result.status, 200);
        const payload = await result.json();
        assert.equal(payload.cloudName, env.CLOUDINARY_CLOUD_NAME);
        assert.equal(payload.params.overwrite, false);
        assert.equal(JSON.stringify(payload).includes(env.CLOUDINARY_API_SECRET), false);
        const denied = await fetch(`${origin}/api/cloudinary-signature`, {
            method: "POST", headers: { Origin: "https://evil.test", Authorization: `Bearer ${token}` },
        });
        assert.equal(denied.status, 403);
        assert.equal((await fetch(`${origin}/api/cloudinary-signature`)).status, 405);
        assert.equal((await fetch(`${origin}/api/cloudinary-signature`, { method: "POST" })).status, 401);
    } finally {
        child.kill("SIGTERM");
        await exit;
    }
});
