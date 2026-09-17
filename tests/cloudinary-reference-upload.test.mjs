import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";
import {
    CLOUDINARY_MAX_BYTES, CLOUDINARY_TARGET_BYTES, CLOUDINARY_MAX_PIXELS,
    usesCloudinaryReferenceHost, isCloudinaryImageUrl, isLegacyImgBbImageUrl,
    prepareCloudinaryReferenceBlob, uploadCloudinaryReferenceBlob, cloudinaryReferenceSource,
    cloudinaryCredentials, signCloudinaryUpload,
} from "../cloudinary-reference-upload.js";
import { inlineTudouGeminiReferences } from "../tudou-reference-images.js";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const config = {
    provider: "tudou", cloudinaryCloudName: "test-cloud", cloudinaryApiKey: "123456",
    cloudinaryApiSecret: "synthetic-user-secret",
};
const input = new Blob(["original-image"], { type: "image/png" });
const decodeImage = async () => ({ image: {}, width: 4096, height: 4096, dispose() {} });
function responseImage(body, cloudName = config.cloudinaryCloudName) {
    const publicId = body.get("public_id");
    return {
        public_id: publicId,
        secure_url: `https://res.cloudinary.com/${cloudName}/image/upload/v1/${publicId}.png`,
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

test("personal upload signs locally and sends one direct request without any API Secret", async () => {
    const calls = [], progress = [];
    const url = await uploadCloudinaryReferenceBlob({ ...config, apiKey: "provider-secret" }, input, {
        decodeImage, onProgress: value => progress.push(value),
        fetchImpl: async (url, init) => {
            calls.push({ url, init });
            assert.equal(url, "https://api.cloudinary.com/v1_1/test-cloud/image/upload");
            assert.equal(init.headers, undefined);
            assert.equal(init.credentials, "omit");
            const body = init.body;
            assert.equal(body.get("api_key"), "123456");
            const serialized = ["overwrite", "public_id", "timestamp"].map(key => `${key}=${body.get(key)}`).join("&");
            assert.equal(body.get("signature"), createHash("sha256").update(serialized + config.cloudinaryApiSecret).digest("hex"));
            assert.equal(body.get("overwrite"), "false");
            assert.deepEqual(await body.get("file").arrayBuffer(), await input.arrayBuffer());
            assert.equal(body.get("file").name, "reference.png");
            assert.equal(body.has("api_secret"), false);
            assert.equal(body.has("Authorization"), false);
            assert.equal(body.has("upload_preset"), false);
            assert.ok(!Array.from(body.values()).includes(config.cloudinaryApiSecret));
            assert.ok(!Array.from(body.values()).includes("provider-secret"));
            return Response.json(responseImage(body));
        },
    });
    assert.equal(calls.length, 1);
    assert.equal(url, responseImage(calls[0].init.body).secure_url);
    assert.deepEqual(progress.map(value => value.stage), ["preparing", "uploading", "uploading"]);
});

test("each missing personal credential stops before reading pixels or contacting any service", async () => {
    for (const key of ["cloudinaryCloudName", "cloudinaryApiKey", "cloudinaryApiSecret"]) {
        await assert.rejects(uploadCloudinaryReferenceBlob({ ...config, [key]: "" }, input, {
            decodeImage: () => assert.fail("unexpected decode"), fetchImpl: () => assert.fail("unexpected request"),
        }), /你自己的 Cloudinary/);
    }
    await assert.rejects(uploadCloudinaryReferenceBlob({ cloudinaryUploadToken: "old-access-code" }, input), /你自己的 Cloudinary/);
});

test("Cloudinary failures do not fall back to ImgBB or submit a model request", async () => {
    let calls = 0;
    await assert.rejects(uploadCloudinaryReferenceBlob(config, input, {
        decodeImage, fetchImpl: async () => { calls++; return Response.json({ error: { message: "denied" } }, { status: 401 }); },
    }), /denied/);
    assert.equal(calls, 1);
});

test("personal credentials are trimmed and invalid destinations rejected before any request", async () => {
    assert.deepEqual(cloudinaryCredentials({
        ...config, cloudinaryCloudName: " test-cloud ", cloudinaryApiKey: " 123456 ",
    }), { cloudName: "test-cloud", apiKey: "123456", apiSecret: config.cloudinaryApiSecret });
    for (const bad of [{ cloudinaryCloudName: "evil/path" }, { cloudinaryApiKey: "not-a-key" }]) {
        await assert.rejects(uploadCloudinaryReferenceBlob({ ...config, ...bad }, input, {
            decodeImage: () => assert.fail("unexpected decode"),
            fetchImpl: () => assert.fail("unexpected request"),
        }), /格式不正确/);
    }
});

test("upload rejects insecure or unrelated result URLs and mismatched asset identifiers", async () => {
    for (const mutate of [
        result => { result.secure_url = "https://evil.test/reference.png"; },
        result => { result.secure_url = result.secure_url.replace("test-cloud", "other-cloud"); },
        result => { result.public_id = "different-asset"; },
    ]) {
        await assert.rejects(uploadCloudinaryReferenceBlob(config, input, {
            decodeImage, fetchImpl: async (_url, init) => {
                const result = responseImage(init.body);
                mutate(result);
                return Response.json(result);
            },
        }), /有效的参考图链接/);
    }
});

test("stalled uploads time out without an automatic paid retry", async () => {
    let calls = 0;
    await assert.rejects(uploadCloudinaryReferenceBlob(config, input, {
        decodeImage, timeoutMs: 20,
        fetchImpl: async (_url, init) => {
            calls++;
            return new Promise((_resolve, reject) => init.signal.addEventListener("abort", () => reject(init.signal.reason)));
        },
    }), /上传超时/);
    assert.equal(calls, 1);
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
    const reference = Object.freeze({ dataUrl: "blob:preview", storageKey: "original-key" });
    let calls = 0;
    await cloudinaryReferenceSource(config, reference, {
        decodeImage, readStoredBlob: async key => { assert.equal(key, "original-key"); return input; },
        fetchImpl: async (_url, init) => { calls++; return Response.json(responseImage(init.body)); },
    });
    assert.equal(reference.dataUrl, "blob:preview");
    assert.equal(reference.storageKey, "original-key");
    assert.equal(calls, 1);
});

test("browser signer uses SHA-256, unique non-overwriting IDs and no incoming transformations", async () => {
    const data = await signCloudinaryUpload({ ...config, transformation: "w_1", public_id: "victim" });
    const serialized = Object.keys(data.params).sort().map(key => `${key}=${data.params[key]}`).join("&");
    assert.equal(data.signature, createHash("sha256").update(serialized + config.cloudinaryApiSecret).digest("hex"));
    assert.equal(data.params.overwrite, false);
    assert.match(data.params.public_id, /^laowu-reference\/[0-9a-f-]{36}$/);
    assert.ok(Math.abs(data.params.timestamp - Date.now() / 1000) < 5);
    assert.equal(data.params.transformation, undefined);
    assert.equal(JSON.stringify(data).includes(config.cloudinaryApiSecret), false);
    assert.notEqual((await signCloudinaryUpload(config)).params.public_id, data.params.public_id);
});

test("switching accounts never reuses another user's key or destination", async () => {
    for (const personal of [config, { ...config, cloudinaryCloudName: "second-cloud", cloudinaryApiKey: "7890", cloudinaryApiSecret: "second-secret" }]) {
        await uploadCloudinaryReferenceBlob(personal, input, {
            decodeImage, fetchImpl: async (url, init) => {
                assert.equal(url, `https://api.cloudinary.com/v1_1/${personal.cloudinaryCloudName}/image/upload`);
                assert.equal(init.body.get("api_key"), personal.cloudinaryApiKey);
                const serialized = ["overwrite", "public_id", "timestamp"].map(key => `${key}=${init.body.get(key)}`).join("&");
                assert.equal(init.body.get("signature"), createHash("sha256").update(serialized + personal.cloudinaryApiSecret).digest("hex"));
                return Response.json(responseImage(init.body, personal.cloudinaryCloudName));
            },
        });
    }
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

test("settings accept personal Cloudinary credentials instead of shared application access codes", () => {
    const settings = between('className:"api-settings-page"', 'className:"api-site-tabs"');
    for (const key of ["cloudinaryCloudName", "cloudinaryApiKey", "cloudinaryApiSecret"]) assert.ok(settings.includes(key));
    assert.match(settings, /secret\?_o.Password:_o/);
    assert.doesNotMatch(settings, /上传服务访问码|cloudinaryUploadToken|ImgBB|imgbbApiKey|referenceImageHost/);
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

test("neither client nor local server depends on a shared signing endpoint or server credentials", async () => {
    const client = await readFile(new URL("../cloudinary-reference-upload.js", import.meta.url), "utf8");
    const server = await readFile(new URL("../local-preview-server.mjs", import.meta.url), "utf8");
    for (const source of [client, server]) assert.doesNotMatch(source, /cloudinary-signature|CLOUDINARY_API_SECRET|CLOUDINARY_UPLOAD_TOKEN_HASHES/);
    await assert.rejects(readFile(new URL("../api/cloudinary-signature.js", import.meta.url)), { code: "ENOENT" });
});
