export const RUNNINGHUB_ORIGIN = "https://www.runninghub.ai";
export const RUNNINGHUB_LLM_ORIGIN = "https://llm.runninghub.ai";
export const RUNNINGHUB_SITE_ID = "runninghub";
export const RUNNINGHUB_SITE_NAME = "RH";
export const RUNNINGHUB_SITE_MODELS = ["nano-banana-pro", "gpt-image-2.5", "google/gemini-3.7-flash"];
export const RUNNINGHUB_IMAGE_MODELS = ["nano-banana-pro", "gpt-image-2.5"];
export const RUNNINGHUB_TEXT_MODELS = ["google/gemini-3.7-flash"];

export const RUNNINGHUB_GPT25_VARIANTS = ["flare", "sunburst"];
export const RUNNINGHUB_GPT25_MODES = ["fixed", "official"];

export const RUNNINGHUB_ASPECT_RATIOS = [
    "1:1",
    "16:9",
    "9:16",
    "4:3",
    "3:4",
    "3:2",
    "2:3",
    "5:4",
    "4:5",
    "21:9",
];

export const RUNNINGHUB_GPT_IMAGE_ASPECT_RATIOS = [
    "1:1",
    "1:2",
    "2:1",
    "1:3",
    "3:1",
    "2:3",
    "3:2",
    "3:4",
    "4:3",
    "4:5",
    "5:4",
    "9:16",
    "16:9",
    "21:9",
    "9:21",
];

const GPT_IMAGE_TEXT_PRICES = {
    low: { "1k": 0.009, "2k": 0.018, "4k": 0.027 },
    medium: { "1k": 0.054, "2k": 0.108, "4k": 0.162 },
    high: { "1k": 0.198, "2k": 0.396, "4k": 0.594 },
};

const GPT_IMAGE_REFERENCE_PRICES = {
    low: { "1k": 0.027, "2k": 0.054, "4k": 0.081 },
    medium: { "1k": 0.054, "2k": 0.108, "4k": 0.162 },
    high: { "1k": 0.198, "2k": 0.396, "4k": 0.594 },
};

const REFERENCE_TARGET_BYTES = 10 * 1024 * 1024;
const REFERENCE_MAX_EDGE = 2048;
const TASK_TIMEOUT_MS = 10 * 60 * 1000;
const uploadedReferences = new WeakMap();

export function isRunningHubSite(config) {
    return config?.provider === "runninghub"
        || String(config?.baseUrl || "").includes("runninghub.ai");
}

export function runningHubImageModel(config) {
    const value = String(config?.model || config?.imageModel || "nano-banana-pro")
        .trim()
        .toLowerCase()
        .split("::")
        .at(-1);
    if (value === "gpt-image-2" || value === "gpt-image-2-all") return "gpt-image-2";
    if (value === "gpt-image-2.5" || value === "gpt-image-2-5") return "gpt-image-2.5";
    return "nano-banana-pro";
}

export function runningHubResolution(config) {
    const value = String(config?.quality || "auto").trim().toLowerCase();
    if (["1k", "2k", "4k"].includes(value)) return value;
    if (["low", "standard"].includes(value)) return "1k";
    if (["medium", "hd"].includes(value)) return "2k";
    if (value === "high") return "4k";
    return "1k";
}

export function runningHubQuality(config) {
    const value = String(config?.gptImageQuality || "medium").trim().toLowerCase();
    const allowed = runningHubImageModel(config) === "gpt-image-2.5"
        ? ["auto", "low", "medium", "high", "xhigh", "max"]
        : ["low", "medium", "high"];
    return allowed.includes(value) ? value : "medium";
}

export function runningHubGpt25Variant(config) {
    const value = String(config?.runningHubGpt25Variant || "flare")
        .trim()
        .toLowerCase();
    return RUNNINGHUB_GPT25_VARIANTS.includes(value) ? value : "flare";
}

export function runningHubGpt25Mode(config) {
    const value = String(config?.runningHubGpt25Mode || "fixed").trim().toLowerCase();
    return RUNNINGHUB_GPT25_MODES.includes(value) ? value : "fixed";
}

export function runningHubGpt25Background(config) {
    const value = String(config?.runningHubBackground || "auto")
        .trim()
        .toLowerCase();
    return ["auto", "opaque", "transparent"].includes(value) ? value : "auto";
}

export function runningHubGpt25OutputFormat(config) {
    const value = String(config?.runningHubOutputFormat || "png")
        .trim()
        .toLowerCase();
    return ["png", "jpeg", "webp"].includes(value) ? value : "png";
}

export function runningHubSupportedAspectRatios(config) {
    const model = runningHubImageModel(config);
    if (model === "gpt-image-2.5" && runningHubGpt25Mode(config) === "fixed") {
        // Economy requests reject the five extra ratios despite the shared docs.
        return RUNNINGHUB_ASPECT_RATIOS;
    }
    return ["gpt-image-2", "gpt-image-2.5"].includes(model)
        ? RUNNINGHUB_GPT_IMAGE_ASPECT_RATIOS
        : RUNNINGHUB_ASPECT_RATIOS;
}

export function runningHubAspectRatio(config) {
    const value = String(config?.size || "auto").trim().toLowerCase();
    const ratios = runningHubSupportedAspectRatios(config);
    if (ratios.includes(value)) return value;
    return undefined;
}

export function runningHubReferenceAspectRatio(config, width, height) {
    if (runningHubImageModel(config) !== "gpt-image-2.5"
        || String(config?.size || "auto").trim().toLowerCase() !== "auto") {
        return config;
    }
    if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
        throw new Error("RH 无法读取第一张参考图的比例，请重新载入图片或手动选择比例");
    }
    const ratio = width / height;
    // Log distance treats portrait and landscape matches symmetrically.
    const size = runningHubSupportedAspectRatios(config).reduce((best, candidate) => {
        const distance = (value) => {
            const [w, h] = value.split(":").map(Number);
            return Math.abs(Math.log(ratio / (w / h)));
        };
        return distance(candidate) < distance(best) ? candidate : best;
    });
    return { ...config, size };
}

export function runningHubImagePrice(config, hasReferences = false) {
    const resolution = runningHubResolution(config);
    if (runningHubImageModel(config) === "nano-banana-pro") {
        return resolution === "4k" ? 0.07 : 0.06;
    }
    if (runningHubImageModel(config) === "gpt-image-2.5") {
        return runningHubGpt25Mode(config) === "fixed" ? 0.03 : null;
    }
    const prices = hasReferences ? GPT_IMAGE_REFERENCE_PRICES : GPT_IMAGE_TEXT_PRICES;
    return prices[runningHubQuality(config)][resolution];
}

export function runningHubImageRequestSpec(config, prompt, imageUrls = []) {
    const hasReferences = imageUrls.length > 0;
    const aspectRatio = runningHubAspectRatio(config);
    const model = runningHubImageModel(config);
    const isGptImage25 = model === "gpt-image-2.5";
    const isGptImage = model === "gpt-image-2";
    const variant = runningHubGpt25Variant(config);
    const isOfficial25 = isGptImage25 && runningHubGpt25Mode(config) === "official";
    const apiModel = isOfficial25
        ? "g-2.5-official-token"
        : "g-2.5";
    const variantPath = isGptImage25 ? `/${variant}` : "";
    return {
        endpoint: isGptImage25
            ? hasReferences
                ? `/openapi/v2/rhart-image-${apiModel}${variantPath}/${isOfficial25 ? "edit" : "image-to-image"}`
                : `/openapi/v2/rhart-image-${apiModel}${variantPath}/text-to-image`
            : isGptImage
            ? hasReferences
                ? "/openapi/v2/rhart-image-g-2-official/image-to-image"
                : "/openapi/v2/rhart-image-g-2-official/text-to-image"
            : hasReferences
                ? "/openapi/v2/rhart-image-n-pro/edit"
                : "/openapi/v2/rhart-image-n-pro/text-to-image",
        body: {
            ...(hasReferences ? { imageUrls } : {}),
            prompt: String(prompt || "").trim(),
            ...(aspectRatio ? { aspectRatio } : {}),
            resolution: runningHubResolution(config),
            ...(isGptImage
                ? { quality: runningHubQuality(config) }
                : isGptImage25 && isOfficial25
                    ? {
                        background: runningHubGpt25Background(config),
                        quality: runningHubQuality(config),
                        outputFormat: runningHubGpt25OutputFormat(config),
                    }
                    : {}),
        },
    };
}

function abortError(signal) {
    return signal?.reason instanceof Error
        ? signal.reason
        : new DOMException("Aborted", "AbortError");
}

function throwIfAborted(signal) {
    if (signal?.aborted) throw abortError(signal);
}

function wait(ms, signal) {
    return new Promise((resolve, reject) => {
        throwIfAborted(signal);
        const timer = setTimeout(resolve, ms);
        signal?.addEventListener("abort", () => {
            clearTimeout(timer);
            reject(abortError(signal));
        }, { once: true });
    });
}

export function runningHubErrorMessage(payload, fallback = "RH 请求失败") {
    const record = payload?.data && typeof payload.data === "object" ? payload.data : payload;
    const code = String(record?.errorCode || payload?.code || "").trim();
    const rawMessage = String(
        record?.errorMessage
        || payload?.message
        || payload?.msg
        || fallback,
    ).trim();

    if (code === "1014") {
        return "当前 RH 密钥不能调用模型接口，请更换 Enterprise-Shared API Key";
    }
    if (code === "416") return "RH 余额不足，请充值后重试";
    if (["401", "802", "1002"].includes(code)) return "RH API Key 无效或已被禁用";
    if (["421", "1003"].includes(code)) return "RH 并发或频率已达上限，请稍后重试";
    if (["809", "1008"].includes(code)) return "RH 参考图过大，请压缩后重试";
    if (["1501", "1505"].includes(code)) return `RH 内容审核未通过：${rawMessage}`;
    return code ? `${rawMessage}（RH ${code}）` : rawMessage;
}

async function fetchJson(url, init, fetchImpl) {
    const response = await fetchImpl(url, init);
    const text = await response.text();
    let payload = {};
    try {
        payload = text ? JSON.parse(text) : {};
    } catch {
        payload = { message: text || `HTTP ${response.status}` };
    }
    if (!response.ok) {
        const error = new Error(runningHubErrorMessage(payload, `RH 请求失败（HTTP ${response.status}）`));
        error.status = response.status;
        throw error;
    }
    return payload;
}

function authHeaders(apiKey, json = false) {
    return {
        Authorization: `Bearer ${String(apiKey || "").trim()}`,
        ...(json ? { "Content-Type": "application/json" } : {}),
    };
}

function taskRecord(payload) {
    return payload?.data && typeof payload.data === "object" && !Array.isArray(payload.data)
        ? payload.data
        : payload || {};
}

export function parseRunningHubTask(payload) {
    const record = taskRecord(payload);
    const status = String(record.status || "").trim().toUpperCase();
    const taskId = String(record.taskId || record.task_id || "").trim();
    const errorCode = String(record.errorCode || "").trim();

    if (status === "FAILED" || errorCode) {
        return { status: "failed", taskId, error: runningHubErrorMessage(record) };
    }
    if (status === "SUCCESS") {
        const urls = (Array.isArray(record.results) ? record.results : [])
            .map((result) => String(result?.url || "").trim())
            .filter((url) => /^https?:\/\//i.test(url));
        return urls.length
            ? { status: "completed", taskId, urls, usage: record.usage || null }
            : { status: "failed", taskId, error: "RH 任务成功但没有返回图片" };
    }
    if (["QUEUED", "RUNNING"].includes(status) || taskId) {
        return { status: "pending", taskId };
    }
    return { status: "failed", taskId, error: runningHubErrorMessage(record, "RH 返回了未知任务状态") };
}

function retryableQueryError(error) {
    return error?.status === 408
        || error?.status === 425
        || error?.status === 429
        || error?.status >= 500;
}

export async function runRunningHubImageGeneration(config, prompt, imageUrls = [], options = {}) {
    if (!String(config?.apiKey || "").trim()) throw new Error("请先填写 RH API Key");
    if (!String(prompt || "").trim()) throw new Error("请输入提示词");
    const maxReferences = runningHubImageModel(config) === "gpt-image-2.5" && runningHubGpt25Mode(config) === "official"
        ? 16
        : 10;
    if (imageUrls.length > maxReferences) throw new Error(`RH 最多支持 ${maxReferences} 张参考图`);

    const fetchImpl = options.fetchImpl || fetch;
    const sleep = options.sleep || wait;
    const now = options.now || Date.now;
    const startedAt = now();
    const spec = runningHubImageRequestSpec(config, prompt, imageUrls);
    const submitted = await fetchJson(`${RUNNINGHUB_ORIGIN}${spec.endpoint}`, {
        method: "POST",
        headers: authHeaders(config.apiKey, true),
        body: JSON.stringify(spec.body),
        signal: options.signal,
    }, fetchImpl);

    let parsed = parseRunningHubTask(submitted);
    if (parsed.status === "completed") return parsed.urls;
    if (parsed.status === "failed") throw new Error(parsed.error);
    if (!parsed.taskId) throw new Error("RH 没有返回 taskId");

    await sleep(2_000, options.signal);
    let consecutiveErrors = 0;
    while (now() - startedAt < (options.timeoutMs || TASK_TIMEOUT_MS)) {
        throwIfAborted(options.signal);
        try {
            const result = await fetchJson(`${RUNNINGHUB_ORIGIN}/openapi/v2/query`, {
                method: "POST",
                headers: authHeaders(config.apiKey, true),
                body: JSON.stringify({ taskId: parsed.taskId }),
                signal: options.signal,
            }, fetchImpl);
            consecutiveErrors = 0;
            parsed = parseRunningHubTask(result);
            if (parsed.status === "completed") return parsed.urls;
            if (parsed.status === "failed") throw new Error(parsed.error);
        } catch (error) {
            if (options.signal?.aborted || !retryableQueryError(error) || ++consecutiveErrors > 3) throw error;
        }
        await sleep(consecutiveErrors ? Math.min(10_000, 3_000 * (consecutiveErrors + 1)) : 3_000, options.signal);
    }
    throw new Error("RH 图像生成超时，请稍后重试");
}

async function decodeImage(blob) {
    if (typeof createImageBitmap === "function") {
        try {
            const image = await createImageBitmap(blob, { imageOrientation: "from-image" });
            return { image, width: image.width, height: image.height, dispose: () => image.close?.() };
        } catch {
            // Fall through to the browser image element path.
        }
    }
    const url = URL.createObjectURL(blob);
    try {
        const image = await new Promise((resolve, reject) => {
            const element = new Image();
            element.onload = () => resolve(element);
            element.onerror = () => reject(new Error("RH 参考图解码失败"));
            element.src = url;
        });
        return {
            image,
            width: image.naturalWidth || image.width,
            height: image.naturalHeight || image.height,
            dispose: () => URL.revokeObjectURL(url),
        };
    } catch (error) {
        URL.revokeObjectURL(url);
        throw error;
    }
}

async function encodeJpeg(image, width, height, quality) {
    const canvas = typeof OffscreenCanvas === "function"
        ? new OffscreenCanvas(width, height)
        : document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error("RH 参考图压缩失败");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);
    return typeof canvas.convertToBlob === "function"
        ? canvas.convertToBlob({ type: "image/jpeg", quality })
        : new Promise((resolve, reject) => canvas.toBlob(
            (blob) => blob ? resolve(blob) : reject(new Error("RH 参考图压缩失败")),
            "image/jpeg",
            quality,
        ));
}

export async function prepareRunningHubReferenceBlob(blob, signal) {
    throwIfAborted(signal);
    if (["image/jpeg", "image/png"].includes(blob.type) && blob.size <= REFERENCE_TARGET_BYTES) {
        return blob;
    }

    const decoded = await decodeImage(blob);
    try {
        if (!decoded.width || !decoded.height) throw new Error("RH 参考图尺寸无效");
        let best = null;
        for (const edge of [REFERENCE_MAX_EDGE, 1800, 1600, 1280, 1024]) {
            const scale = Math.min(1, edge / Math.max(decoded.width, decoded.height));
            const width = Math.max(1, Math.round(decoded.width * scale));
            const height = Math.max(1, Math.round(decoded.height * scale));
            for (const quality of [0.92, 0.84, 0.76, 0.68]) {
                throwIfAborted(signal);
                const candidate = await encodeJpeg(decoded.image, width, height, quality);
                if (!best || candidate.size < best.size) best = candidate;
                if (candidate.size <= REFERENCE_TARGET_BYTES) return candidate;
            }
        }
        if (best?.size <= REFERENCE_TARGET_BYTES) return best;
        throw new Error("RH 临时参考图压缩后仍超过 10 MB");
    } finally {
        decoded.dispose();
    }
}

export async function uploadRunningHubReferenceBlob(config, sourceBlob, options = {}) {
    if (!String(config?.apiKey || "").trim()) throw new Error("请先填写 RH API Key");
    const prepared = await prepareRunningHubReferenceBlob(sourceBlob, options.signal);
    let pending = uploadedReferences.get(prepared);
    if (!pending) {
        pending = (async () => {
            const form = new FormData();
            form.set("file", prepared, prepared.type === "image/png" ? "reference.png" : "reference.jpg");
            const payload = await fetchJson(`${RUNNINGHUB_ORIGIN}/openapi/v2/media/upload/binary`, {
                method: "POST",
                headers: authHeaders(config.apiKey),
                body: form,
                signal: options.signal,
            }, options.fetchImpl || fetch);
            if (payload?.code != null && ![0, 200].includes(Number(payload.code))) {
                throw new Error(runningHubErrorMessage(payload, "RH 参考图上传失败"));
            }
            const url = String(payload?.data?.download_url || "").trim();
            if (!/^https?:\/\//i.test(url)) throw new Error("RH 上传成功但没有返回参考图地址");
            return url;
        })();
        uploadedReferences.set(prepared, pending);
        pending.catch(() => uploadedReferences.delete(prepared));
    }
    return pending;
}

export async function fetchRunningHubAccount(config, options = {}) {
    const apiKey = String(config?.apiKey || "").trim();
    if (!apiKey) throw new Error("请先填写 RH API Key");
    const payload = await fetchJson(`${RUNNINGHUB_ORIGIN}/uc/openapi/accountStatus`, {
        method: "POST",
        headers: authHeaders(apiKey, true),
        body: JSON.stringify({ apikey: apiKey }),
        signal: options.signal,
    }, options.fetchImpl || fetch);
    if (Number(payload?.code) !== 0 || !payload?.data) {
        throw new Error(runningHubErrorMessage(payload, "RH 账户信息查询失败"));
    }
    return payload.data;
}
