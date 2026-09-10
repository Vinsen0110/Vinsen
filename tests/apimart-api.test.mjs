import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
    APIMART_BACKEND_MODEL,
    APIMART_GPT_FIXED_BACKEND_MODEL,
    APIMART_GPT_FIXED_QUALITY,
    APIMART_GPT_OFFICIAL_BACKEND_MODEL,
    APIMART_GPT25_DEFAULT_BACKEND_MODEL,
    APIMART_GPT25_SUNBURST_BACKEND_MODEL,
    APIMART_GPT25_FIXED_BACKEND_MODEL,
    APIMART_ORIGIN,
    APIMART_SITE_NAME,
    APIMART_TEXT_MODELS,
    APIMART_UPLOAD_MAX_BYTES,
    apiMartGptBackground,
    apiMartGptOutputFormat,
    apiMartImagePrice,
    apiMartImageRequestSpec,
    fetchApiMartTokenBalance,
    fetchApiMartUserBalance,
    prepareApiMartReferenceBlob,
    runApiMartImageGeneration,
    uploadApiMartReferenceBlob,
} from "../apimart-api.js";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

function jsonResponse(payload, status = 200) {
    return new Response(JSON.stringify(payload), {
        status,
        headers: { "Content-Type": "application/json" },
    });
}

test("APIMart keeps the shared canvas name but requests the standard backend model", () => {
    assert.equal(APIMART_SITE_NAME, "Mart");
    const spec = apiMartImageRequestSpec(
        { model: "apimart::nano-banana-pro", size: "16:9", quality: "2k" },
        "draw a poster",
        ["https://cdn.example/reference.png"],
    );

    assert.equal(APIMART_BACKEND_MODEL, "gemini-3-pro-image-preview");
    assert.equal(spec.endpoint, "/images/generations");
    assert.deepEqual(spec.body, {
        model: "gemini-3-pro-image-preview",
        prompt: "draw a poster",
        size: "16:9",
        resolution: "2K",
        n: 1,
        response_format: "url",
        image_urls: ["https://cdn.example/reference.png"],
    });
    assert.throws(
        () => apiMartImageRequestSpec(
            { model: "nano-banana-pro", size: "1:1", quality: "1k" },
            "draw",
            Array.from({ length: 15 }, (_, index) => `https://cdn.example/${index}.png`),
        ),
        /最多支持 14 张参考图/,
    );
});

test("APIMart converts 0.3 and 0.4 Credits to the displayed dollar balance unit", () => {
    for (const quality of ["auto", "1k", "2k"]) {
        assert.equal(apiMartImagePrice({ quality }), 0.03);
    }
    assert.equal(apiMartImagePrice({ quality: "4k" }), 0.04);
});

test("APIMart GPT Image 2 uses fixed mode without quality and official mode for auto and quality tiers", () => {
    const fixed = apiMartImageRequestSpec(
        {
            model: "apimart::gpt-image-2",
            size: "16:9",
            quality: "2k",
            gptImageQuality: APIMART_GPT_FIXED_QUALITY,
        },
        "draw a product photo",
        ["https://cdn.example/reference.png"],
    );
    assert.deepEqual(fixed.body, {
        model: APIMART_GPT_FIXED_BACKEND_MODEL,
        prompt: "draw a product photo",
        size: "16:9",
        resolution: "2k",
        n: 1,
        image_urls: ["https://cdn.example/reference.png"],
    });
    assert.equal("quality" in fixed.body, false);

    for (const quality of ["auto", "low", "medium", "high"]) {
        const official = apiMartImageRequestSpec(
            {
                model: "gpt-image-2",
                size: "5:4",
                quality: "4k",
                gptImageQuality: quality,
            },
            "draw it",
        );
        assert.equal(official.body.model, APIMART_GPT_OFFICIAL_BACKEND_MODEL);
        assert.equal(official.body.quality, quality);
        assert.equal(official.body.resolution, "4k");
        assert.equal(official.body.output_format, "png");
        assert.equal(official.body.background, "auto");
    }

    const transparent = apiMartImageRequestSpec(
        {
            model: "gpt-image-2",
            size: "16:9",
            quality: "2k",
            gptImageQuality: "medium",
            apimartOutputFormat: "jpeg",
            apimartBackground: "transparent",
        },
        "separate the elements",
    );
    assert.equal(transparent.body.output_format, "png");
    assert.equal(transparent.body.background, "transparent");
    assert.equal(apiMartGptOutputFormat({ apimartOutputFormat: "webp", apimartBackground: "transparent" }), "webp");
    assert.equal(apiMartGptBackground({ apimartBackground: "opaque" }), "opaque");

    assert.equal(apiMartImagePrice({ model: "gpt-image-2", quality: "1k", gptImageQuality: "fixed" }), 0.0085);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2", quality: "2k", gptImageQuality: "fixed" }), 0.014);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2", quality: "4k", gptImageQuality: "fixed" }), 0.021);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2", size: "1:1", quality: "2k", gptImageQuality: "low" }), null);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2", size: "16:9", quality: "2k", gptImageQuality: "medium" }), null);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2", size: "auto", quality: "4k", gptImageQuality: "high" }), null);
});

test("APIMart GPT Image 2.5 sends the selected backend variant with its extended quality tiers", () => {
    for (const quality of ["auto", "low", "medium", "high", "xhigh", "max"]) {
        const spec = apiMartImageRequestSpec(
            {
                model: "gpt-image-2.5",
                size: "16:9",
                quality: "2k",
                gptImageQuality: quality,
            },
            "draw it",
        );
        assert.equal(spec.body.model, APIMART_GPT25_DEFAULT_BACKEND_MODEL);
        assert.equal(spec.body.quality, quality);
        assert.equal(spec.body.resolution, "2k");
    }
    const sunburst = apiMartImageRequestSpec(
        { model: "gpt-image-2.5", apimartGpt25Variant: "sunburst", quality: "2k", gptImageQuality: "xhigh" },
        "edit it",
    );
    assert.equal(sunburst.body.model, APIMART_GPT25_SUNBURST_BACKEND_MODEL);
});

test("APIMart GPT Image 2.5 fixed mode uses ext with the selected version and fixed price", () => {
    const spec = apiMartImageRequestSpec(
        {
            model: "gpt-image-2.5",
            apimartGpt25Variant: "sunburst",
            gptImageQuality: "fixed",
            quality: "2k",
            size: "auto",
        },
        "edit it",
    );
    assert.equal(spec.body.model, APIMART_GPT25_FIXED_BACKEND_MODEL);
    assert.equal(spec.body.version, "sunburst");
    assert.equal(spec.body.resolution, "2k");
    assert.equal(spec.body.size, "auto");
    assert.equal("quality" in spec.body, false);
    assert.equal("output_format" in spec.body, false);
    assert.equal("background" in spec.body, false);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2.5", gptImageQuality: "fixed", quality: "1k" }), 0.0085);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2.5", gptImageQuality: "fixed", quality: "2k" }), 0.014);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2.5", gptImageQuality: "fixed", quality: "4k" }), 0.021);
    assert.equal(apiMartImagePrice({ model: "gpt-image-2.5", gptImageQuality: "high", quality: "2k" }), null);
});

test("APIMart submits once and always polls the original task id", async () => {
    const calls = [];
    const responses = [
        { data: [{ task_id: "task-original", status: "pending" }] },
        { data: { id: "internal-record", status: "processing" } },
        {
            data: {
                id: "another-record",
                status: "completed",
                result: { images: [{ url: ["https://cdn.example/result.png"] }] },
            },
        },
    ];
    let clock = 0;
    const urls = await runApiMartImageGeneration(
        {
            baseUrl: APIMART_ORIGIN,
            apiKey: "apimart-key",
            model: "nano-banana-pro",
            size: "1:1",
            quality: "4k",
        },
        "draw it",
        [],
        {
            fetchImpl: async (url, init) => {
                calls.push({ url: String(url), init });
                return jsonResponse(responses.shift());
            },
            sleep: async () => {},
            now: () => (clock += 1_000),
        },
    );

    assert.deepEqual(urls, ["https://cdn.example/result.png"]);
    assert.equal(calls.filter(({ init }) => init.method === "POST").length, 1);
    assert.equal(calls[0].url, `${APIMART_ORIGIN}/v1/images/generations`);
    assert.deepEqual(JSON.parse(calls[0].init.body), {
        model: "gemini-3-pro-image-preview",
        prompt: "draw it",
        size: "1:1",
        resolution: "4K",
        n: 1,
        response_format: "url",
    });
    assert.equal(calls[1].url, `${APIMART_ORIGIN}/v1/tasks/task-original?language=zh`);
    assert.equal(calls[2].url, `${APIMART_ORIGIN}/v1/tasks/task-original?language=zh`);
});

test("APIMart upload keeps small originals and compresses only an oversized request copy", async () => {
    const small = new Blob([Uint8Array.from([1, 2, 3])], { type: "image/png" });
    assert.equal(await prepareApiMartReferenceBlob(small), small);

    const original = { size: APIMART_UPLOAD_MAX_BYTES + 1, type: "image/png" };
    const compressed = new Blob([Uint8Array.from([1, 2])], { type: "image/webp" });
    const prepared = await prepareApiMartReferenceBlob(original, {
        decodeImage: async () => ({ image: {}, width: 4096, height: 4096, dispose() {} }),
        encodeImage: async () => compressed,
    });
    assert.equal(prepared, compressed);
    assert.equal(original.size, APIMART_UPLOAD_MAX_BYTES + 1);

    const calls = [];
    const uploadedUrl = await uploadApiMartReferenceBlob(
        { baseUrl: APIMART_ORIGIN, apiKey: "apimart-key" },
        small,
        {
            fetchImpl: async (url, init) => {
                calls.push({ url: String(url), init });
                return jsonResponse({ data: { url: "https://cdn.example/upload.png" } });
            },
        },
    );
    assert.equal(uploadedUrl, "https://cdn.example/upload.png");
    assert.equal(calls[0].url, `${APIMART_ORIGIN}/v1/uploads/images`);
    assert.equal(calls[0].init.method, "POST");
    assert.equal(calls[0].init.headers.Authorization, "Bearer apimart-key");
    assert.deepEqual(Array.from(calls[0].init.body.keys()), ["file"]);
});

test("APIMart token and user balances use their documented endpoints", async () => {
    const calls = [];
    const fetchImpl = async (url, init) => {
        calls.push({ url: String(url), init });
        return jsonResponse({ data: { remain_balance: calls.length === 1 ? 12.34 : 56.78 } });
    };

    assert.equal(await fetchApiMartTokenBalance(
        { baseUrl: APIMART_ORIGIN, apiKey: "apimart-key" },
        { fetchImpl },
    ), 12.34);
    assert.equal(await fetchApiMartUserBalance(
        { baseUrl: APIMART_ORIGIN, apiKey: "apimart-key" },
        { fetchImpl },
    ), 56.78);
    assert.equal(calls[0].url, `${APIMART_ORIGIN}/v1/balance`);
    assert.equal(calls[1].url, `${APIMART_ORIGIN}/v1/user/balance`);
    assert.equal(calls[0].init.headers.Authorization, "Bearer apimart-key");
});

test("compiled app keeps APIMart isolated from the four existing sites", () => {
    assert.match(bundle, /provider:"apimart",models:APIMART_SITE_MODELS/);
    assert.match(
        bundle,
        /e\?\.provider==="apimart"\?"apimart":"apilio"/,
        "APIMart must survive provider normalization instead of falling back to Apilio",
    );
    assert.match(bundle, /apiKey:u\?\.apiKey\|\|"",apiKeys:u\?\.apiKeys,activeKeyId:u\?\.activeKeyId,apiFormat:"openai",provider:"apimart"/);
    assert.match(bundle, /isApiMartSite\(r\).*runApiMartImageGeneration\(r,t,\[\]/);
    assert.match(bundle, /isApiMartSite\(a\).*apiMartReferenceSource/);
    assert.match(bundle, /u=isApiMartSite\(e\)\|\|isApiMartSite\(r\).*u\?apiMartImagePrice\(e\)/);
    assert.match(bundle, /isApiMartSite\(K\).*fetchApiMartTokenBalance\(K\)/);
    assert.match(bundle, /isApiMartSite\(le\).*fetchApiMartUserBalance\(le\)/);
    assert.deepEqual(APIMART_TEXT_MODELS, []);
});

test("APIMart settings share multi-key controls without a billing panel", () => {
    assert.match(bundle, /onClick:addSiteKey,children:"\\u6DFB\\u52A0\\u5BC6\\u94A5"/);
    assert.match(bundle, /onClick:deleteSiteKey,children:"\\u5220\\u9664"/);
    assert.match(bundle, /children:siteKeys\.map/);
    assert.doesNotMatch(bundle, /g\?\.provider==="apimart"\?\[/);
    assert.doesNotMatch(bundle, /Nano Banana Pro \\u56FA\\u5B9A\\u6263\\u8D39/);
    assert.match(
        bundle,
        /g\?\.provider==="apilio"\?y\.jsxs\("div",\{className:"api-billing-panel"/,
        "Apilio keeps its billing-group controls",
    );
    assert.match(bundle, /children:"GPT Image 2 \u8BA1\u8D39\u7EC4"/);
    assert.match(bundle, /children:"\u91CD\u65B0\u68C0\u6D4B"/);
    assert.doesNotMatch(
        bundle,
        /g\?\.provider==="apimart"\|\|g\?\.provider==="grsai"/,
        "APIMart must not share the Apilio or Grsai account-balance credential form",
    );
});

test("APIMart GPT Image 2 exposes isolated fixed and official controls", () => {
    assert.match(bundle, /APIMART_GPT_MODE_OPTIONS=\[\{value:"fixed",label:"固定"\},\{value:"official",label:"官方"\}\]/);
    assert.match(bundle, /APIMART_GPT25_VARIANT_OPTIONS=\[\{value:"flare",label:"Flare"\},\{value:"sunburst",label:"Sunburst"\}\]/);
    assert.match(bundle, /APIMART_GPT_OFFICIAL_QUALITY_OPTIONS=\[\{value:"auto",label:"自动"\},\.\.\.GPT_IMAGE_QUALITY_OPTIONS\]/);
    assert.match(bundle, /APIMART_GPT_OUTPUT_FORMAT_OPTIONS=\[\{value:"png",label:"PNG"\},\{value:"jpeg",label:"JPEG"\},\{value:"webp",label:"WebP"\}\]/);
    assert.match(bundle, /APIMART_GPT_BACKGROUND_OPTIONS=\[\{value:"auto",label:"自动"\},\{value:"opaque",label:"不透明"\},\{value:"transparent",label:"透明"\}\]/);
    assert.match(bundle, /\(pr\(e\?\.model\|\|e\?\.imageModel\)==="gpt-image-2\.5"\?\["fixed","auto","low","medium","high","xhigh","max"\]:\["fixed","low","medium","high"\]\)\.includes\(r\)/);
    assert.match(bundle, /isApiMartGptImageConfig\?y\.jsxs\(y\.Fragment/);
    assert.match(bundle, /\(apiMartModeValue==="official"\)\?y\.jsxs\(y\.Fragment/);
    assert.match(bundle, /options:APIMART_GPT_OUTPUT_FORMAT_OPTIONS\.map/);
    assert.match(bundle, /apiMartBackgroundValue==="transparent"&&\$\.value==="jpeg"\?\{disabled:!0\}/);
    assert.match(bundle, /\$==="transparent"&&apiMartOutputValue==="jpeg"\?\{apimartOutputFormat:"png"\}/);
    assert.match(bundle, /gptImageQuality:"fixed"/);
    assert.match(bundle, /hideCredits=isApiMartGptImageModel&&apiMartModeValue!=="fixed"/);
    assert.match(bundle, /isApiMartGptImageModel&&apiMartModeValue==="official"&&"is-apimart-official"/);
    assert.match(bundle, /value:apiMartModeValue,items:APIMART_GPT_MODE_OPTIONS/);
    assert.match(bundle, /value:apiMart25VariantValue,items:APIMART_GPT25_VARIANT_OPTIONS/);
    assert.match(bundle, /widthClass:"w-\[72px\]"/);
    assert.match(bundle, /value:apiMartOutputValue,items:apiMartOutputOptions/);
    assert.match(bundle, /value:apiMartBackgroundValue,items:APIMART_GPT_BACKGROUND_OPTIONS/);
    assert.match(bundle, /disabled:!!g\.disabled,title:g\.title/);
    assert.match(bundle, /s\?null:n==null\?y\.jsx\("span"/);
    const modeOptions = bundle.match(/APIMART_GPT_MODE_OPTIONS=(\[[^\]]+\])/)?.[1] || "";
    assert.doesNotMatch(modeOptions, /value:"auto"/);
});

test("APIMart official output settings survive generation, retry, and project restore", () => {
    assert.match(bundle, /function X0\([^)]*\).*apimartOutputFormat:t\.apimartOutputFormat,apimartBackground:t\.apimartBackground/);
    assert.match(bundle, /normalizeImageModelParams\(e\).*apimartGpt25Variant:u/);
    assert.match(bundle, /imageModelParamsFromConfig\(e\).*apimartGpt25Variant:e\.apimartGpt25Variant/);
    for (const functionName of ["Q0", "vke", "cPe"]) {
        const start = bundle.indexOf(`function ${functionName}(`);
        assert.notEqual(start, -1, `${functionName} must exist`);
        const body = bundle.slice(start, start + 1800);
        assert.match(body, /apimartOutputFormat:t\??\.metadata\?\.apimartOutputFormat\|\|e\.apimartOutputFormat\|\|an\.apimartOutputFormat/);
        assert.match(body, /apimartBackground:t\??\.metadata\?\.apimartBackground\|\|e\.apimartBackground\|\|an\.apimartBackground/);
        assert.match(body, /apimartGpt25Variant:t\??\.metadata\?\.apimartGpt25Variant\|\|e\.apimartGpt25Variant\|\|an\.apimartGpt25Variant/);
    }
    assert.match(bundle, /qM=\{[^}]*model:\{type:"string"\}.*apimartOutputFormat:\{type:"string"\},apimartBackground:\{type:"string"\}/);
    assert.match(bundle, /projectImageModelParams\(e\).*apimartOutputFormat:t\.apimartOutputFormat,apimartBackground:t\.apimartBackground/);
});
