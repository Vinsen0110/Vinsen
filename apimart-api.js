export const APIMART_ORIGIN = "https://api.apimart.ai";
export const APIMART_SITE_ID = "apimart";
export const APIMART_SITE_NAME = "Mart";
export const APIMART_IMAGE_MODELS = ["nano-banana-pro", "gpt-image-2", "gpt-image-2.5"];
export const APIMART_TEXT_MODELS = ["gemini-3.8-flash"];
export const APIMART_SITE_MODELS = [...APIMART_IMAGE_MODELS, ...APIMART_TEXT_MODELS];
export const APIMART_BACKEND_MODEL = "gemini-3-pro-image-preview";
export const APIMART_GPT_FIXED_BACKEND_MODEL = "gpt-image-2";
export const APIMART_GPT_OFFICIAL_BACKEND_MODEL = "gpt-image-2-official";
export const APIMART_GPT25_DEFAULT_BACKEND_MODEL = "gpt-image-2.5-flare";
export const APIMART_GPT25_SUNBURST_BACKEND_MODEL = "gpt-image-2.5-sunburst";
export const APIMART_GPT25_FIXED_BACKEND_MODEL = "gpt-image-2.5-ext";
export const APIMART_GPT_FIXED_QUALITY = "fixed";
export const APIMART_GPT_DEFAULT_OUTPUT_FORMAT = "png";
export const APIMART_GPT_DEFAULT_BACKGROUND = "auto";
export const APIMART_UPLOAD_MAX_BYTES = 20 * 1024 * 1024;
export const APIMART_UPLOAD_TARGET_BYTES = 18 * 1024 * 1024;

const POLL_INTERVAL_MS = 2500;
const TASK_TIMEOUT_MS = 10 * 60 * 1000;
const UPLOAD_TIMEOUT_MS = 3 * 60 * 1000;
const NANO_MAX_REFERENCE_IMAGES = 14;
const GPT_MAX_REFERENCE_IMAGES = 15;
const IMAGE_TYPES = new Set(["image/jpeg", "image/png", "image/webp", "image/gif"]);
const REFERENCE_QUALITIES = [0.92, 0.84, 0.76, 0.68, 0.58, 0.48, 0.38, 0.3];
const REFERENCE_SCALES = [1, 0.9, 0.8, 0.7, 0.6, 0.5, 0.42, 0.35, 0.3, 0.25];
const NANO_ASPECT_RATIOS = new Set([
    "auto", "1:1", "2:3", "3:2", "3:4", "4:3",
    "4:5", "5:4", "9:16", "16:9", "21:9",
]);
const GPT_ASPECT_RATIOS = new Set([
    "auto", "1:1", "2:3", "3:2", "3:4", "4:3",
    "4:5", "5:4", "9:16", "16:9", "2:1", "1:2",
    "3:1", "1:3", "21:9", "9:21",
]);

const GPT_FIXED_PRICE_USD = {
    "1k": 0.0085,
    "2k": 0.014,
    "4k": 0.021,
};

const GPT25_FIXED_PRICE_USD = {
    "1k": 0.0085,
    "2k": 0.014,
    "4k": 0.021,
};

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

function apiRoot(baseUrl) {
    const root = String(baseUrl || APIMART_ORIGIN).trim().replace(/\/+$/, "");
    return /\/v1$/i.test(root) ? root : `${root}/v1`;
}

function authHeaders(apiKey, contentType = "application/json") {
    return {
        Accept: "application/json",
        Authorization: `Bearer ${String(apiKey || "").trim()}`,
        ...(contentType ? { "Content-Type": contentType } : {}),
    };
}

function modelName(config) {
    const value = String(config?.model || config?.imageModel || "nano-banana-pro");
    return value.includes("::") ? value.slice(value.indexOf("::") + 2) : value;
}

export function isApiMartSite(config) {
    return config?.provider === APIMART_SITE_ID
        || /(^|\.)apimart\.ai$/i.test(String(config?.baseUrl || "")
            .replace(/^https?:\/\//i, "")
            .split("/")[0]);
}

export function apiMartResolution(config) {
    const value = String(config?.quality || "auto").trim().toLowerCase();
    if (["4k", "high"].includes(value)) return "4K";
    if (["2k", "medium", "hd"].includes(value)) return "2K";
    return "1K";
}

function apiMartResolutionKey(config) {
    return apiMartResolution(config).toLowerCase();
}

export function apiMartGptImageQuality(config) {
    const value = String(config?.gptImageQuality || APIMART_GPT_FIXED_QUALITY).trim().toLowerCase();
    const is25 = modelName(config) === "gpt-image-2.5";
    const allowed = is25 ? ["fixed", "auto", "low", "medium", "high", "xhigh", "max"] : ["auto", "low", "medium", "high"];
    return allowed.includes(value) ? value : (is25 ? "auto" : APIMART_GPT_FIXED_QUALITY);
}

export function apiMartGpt25BackendModel(config) {
    return String(config?.apimartGpt25Variant || "flare").trim().toLowerCase() === "sunburst"
        ? APIMART_GPT25_SUNBURST_BACKEND_MODEL
        : APIMART_GPT25_DEFAULT_BACKEND_MODEL;
}

export function apiMartGptBackground(config) {
    const value = String(config?.apimartBackground || APIMART_GPT_DEFAULT_BACKGROUND).trim().toLowerCase();
    return ["auto", "opaque", "transparent"].includes(value)
        ? value
        : APIMART_GPT_DEFAULT_BACKGROUND;
}

export function apiMartGptOutputFormat(config) {
    const value = String(config?.apimartOutputFormat || APIMART_GPT_DEFAULT_OUTPUT_FORMAT).trim().toLowerCase();
    const normalized = ["png", "jpeg", "webp"].includes(value)
        ? value
        : APIMART_GPT_DEFAULT_OUTPUT_FORMAT;
    return apiMartGptBackground(config) === "transparent" && normalized === "jpeg"
        ? APIMART_GPT_DEFAULT_OUTPUT_FORMAT
        : normalized;
}

export function apiMartImagePrice(config) {
    const model = modelName(config);
    if (model === "gpt-image-2.5") {
        return apiMartGptImageQuality(config) === APIMART_GPT_FIXED_QUALITY
            ? GPT25_FIXED_PRICE_USD[apiMartResolutionKey(config)]
            : null;
    }
    if (model !== "gpt-image-2") {
        return apiMartResolution(config) === "4K" ? 0.04 : 0.03;
    }
    const resolution = apiMartResolutionKey(config);
    const quality = apiMartGptImageQuality(config);
    if (quality === APIMART_GPT_FIXED_QUALITY) return GPT_FIXED_PRICE_USD[resolution];
    // Official-channel charges depend on actual token usage, especially for image inputs.
    // Do not display a misleading pre-generation estimate in the canvas.
    return null;
}

export function apiMartImageRequestSpec(config, prompt, imageUrls = []) {
    const model = modelName(config);
    if (!["nano-banana-pro", "gpt-image-2", "gpt-image-2.5"].includes(model)) {
        throw new Error("APIMart 当前只支持 Nano Banana Pro、GPT Image 2 和 GPT Image 2.5");
    }
    const maxReferences = model === "gpt-image-2" || model === "gpt-image-2.5" ? 16 : NANO_MAX_REFERENCE_IMAGES;
    if (imageUrls.length > maxReferences) throw new Error(`APIMart ${model === "gpt-image-2" ? "GPT Image 2" : model === "gpt-image-2.5" ? "GPT Image 2.5" : "Nano Banana Pro"} 最多支持 ${maxReferences} 张参考图`);
    const size = String(config?.size || "auto").trim().toLowerCase();
    const supportedRatios = model === "gpt-image-2" || model === "gpt-image-2.5" ? GPT_ASPECT_RATIOS : NANO_ASPECT_RATIOS;
    if (!supportedRatios.has(size) && !/^\d+x\d+$/.test(size)) throw new Error(`APIMart 不支持尺寸比例 ${size}`);

    if (model === "gpt-image-2" || model === "gpt-image-2.5") {
        const quality = apiMartGptImageQuality(config);
        const is25 = model === "gpt-image-2.5";
        const isFixed25 = is25 && quality === APIMART_GPT_FIXED_QUALITY;
        return {
            endpoint: "/images/generations",
            body: {
                model: is25
                    ? (isFixed25 ? APIMART_GPT25_FIXED_BACKEND_MODEL : apiMartGpt25BackendModel(config))
                    : quality === APIMART_GPT_FIXED_QUALITY
                        ? APIMART_GPT_FIXED_BACKEND_MODEL
                        : APIMART_GPT_OFFICIAL_BACKEND_MODEL,
                prompt: String(prompt || "").trim(),
                size,
                resolution: apiMartResolutionKey(config),
                n: 1,
                ...(isFixed25 ? {
                    version: String(config?.apimartGpt25Variant || "flare").trim().toLowerCase() === "sunburst" ? "sunburst" : "flare",
                } : is25 || quality !== APIMART_GPT_FIXED_QUALITY ? {
                    quality,
                    output_format: apiMartGptOutputFormat(config),
                    background: apiMartGptBackground(config),
                } : {}),
                ...(imageUrls.length ? { image_urls: imageUrls } : {}),
            },
        };
    }

    return {
        endpoint: "/images/generations",
        body: {
            model: APIMART_BACKEND_MODEL,
            prompt: String(prompt || "").trim(),
            size,
            resolution: apiMartResolution(config),
            n: 1,
            response_format: "url",
            ...(imageUrls.length ? { image_urls: imageUrls } : {}),
        },
    };
}

function responseError(payload, fallback) {
    return String(
        payload?.error?.message
        || payload?.data?.error?.message
        || payload?.message
        || payload?.msg
        || payload?.detail
        || fallback,
    ).trim();
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
        const error = new Error(responseError(payload, `APIMart 请求失败（HTTP ${response.status}）`));
        error.status = response.status;
        throw error;
    }
    return payload;
}

function taskRecord(payload) {
    if (Array.isArray(payload?.data)) return payload.data[0] || {};
    if (payload?.data && typeof payload.data === "object") return payload.data;
    return payload || {};
}

function collectTaskResultUrls(record) {
    const images = Array.isArray(record?.result?.images)
        ? record.result.images
        : Array.isArray(record?.images)
            ? record.images
            : [];
    const candidates = [
        record?.result?.url,
        record?.url,
        ...images.flatMap((item) => Array.isArray(item?.url) ? item.url : [item?.url]),
    ];
    return Array.from(new Set(candidates
        .map((item) => String(item || "").trim())
        .filter((item) => /^https?:\/\//i.test(item))));
}

export function parseApiMartTask(payload) {
    const record = taskRecord(payload);
    const taskId = String(record.task_id || payload?.task_id || record.id || payload?.id || "").trim();
    const status = String(record.status || record.state || record.task_status || "").trim().toLowerCase();
    const progressValue = Number(record.progress ?? payload?.progress);
    const progress = Number.isFinite(progressValue) ? progressValue : undefined;

    if (["failed", "error", "cancelled", "canceled"].includes(status)) {
        return { status: "failed", taskId, progress, error: responseError(record, "APIMart 任务失败") };
    }
    if (["completed", "succeeded", "success"].includes(status)) {
        const urls = collectTaskResultUrls(record);
        return urls.length
            ? { status: "completed", taskId, progress: 100, urls }
            : { status: "failed", taskId, progress: 100, error: "APIMart 任务完成但没有返回图片" };
    }
    if (["pending", "processing", "queued", "submitted", "running"].includes(status) || taskId) {
        return { status: "pending", taskId, progress };
    }
    return { status: "failed", taskId, progress, error: responseError(record, "APIMart 返回了未知任务状态") };
}

function retryableQueryError(error) {
    return error?.status === 408
        || error?.status === 425
        || error?.status === 429
        || error?.status >= 500;
}

export async function runApiMartImageGeneration(config, prompt, imageUrls = [], options = {}) {
    const apiKey = String(config?.apiKey || "").trim();
    if (!apiKey) throw new Error("请先填写 APIMart API Key");
    if (!String(prompt || "").trim()) throw new Error("请输入提示词");

    const fetchImpl = options.fetchImpl || fetch;
    const sleep = options.sleep || wait;
    const now = options.now || Date.now;
    const startedAt = now();
    const spec = apiMartImageRequestSpec(config, prompt, imageUrls);
    const submitted = await fetchJson(`${apiRoot(config?.baseUrl)}${spec.endpoint}`, {
        method: "POST",
        headers: authHeaders(apiKey),
        body: JSON.stringify(spec.body),
        signal: options.signal,
    }, fetchImpl);
    const first = parseApiMartTask(submitted);
    if (first.status === "completed") return first.urls;
    if (first.status === "failed") throw new Error(first.error);
    if (!first.taskId) throw new Error("APIMart 没有返回 task_id");

    const taskId = first.taskId;
    options.onProgress?.({ stage: "generating", progress: first.progress, taskId });
    await sleep(1500, options.signal);
    let consecutiveErrors = 0;
    while (now() - startedAt < (options.timeoutMs || TASK_TIMEOUT_MS)) {
        throwIfAborted(options.signal);
        try {
            const payload = await fetchJson(
                `${apiRoot(config?.baseUrl)}/tasks/${encodeURIComponent(taskId)}?language=zh`,
                { method: "GET", headers: authHeaders(apiKey), signal: options.signal },
                fetchImpl,
            );
            consecutiveErrors = 0;
            const result = parseApiMartTask(payload);
            options.onProgress?.({ stage: "generating", progress: result.progress, taskId });
            if (result.status === "completed") return result.urls;
            if (result.status === "failed") throw new Error(result.error);
        } catch (error) {
            if (!retryableQueryError(error) || consecutiveErrors >= 2) throw error;
            consecutiveErrors += 1;
        }
        await sleep(options.pollIntervalMs || POLL_INTERVAL_MS, options.signal);
    }
    throw new Error("APIMart 生成超时，请稍后到站点后台查看任务");
}

async function decodeReferenceBlob(blob) {
    if (typeof createImageBitmap === "function") {
        try {
            const image = await createImageBitmap(blob, { imageOrientation: "from-image" });
            return { image, width: image.width, height: image.height, dispose: () => image.close?.() };
        } catch {}
    }
    const objectUrl = URL.createObjectURL(blob);
    try {
        const image = await new Promise((resolve, reject) => {
            const element = new Image();
            element.onload = () => resolve(element);
            element.onerror = () => reject(new Error("APIMart 参考图解码失败"));
            element.src = objectUrl;
        });
        return {
            image,
            width: image.naturalWidth || image.width,
            height: image.naturalHeight || image.height,
            dispose: () => URL.revokeObjectURL(objectUrl),
        };
    } catch (error) {
        URL.revokeObjectURL(objectUrl);
        throw error;
    }
}

async function canvasBlob(canvas, type, quality) {
    if (typeof canvas.convertToBlob === "function") return canvas.convertToBlob({ type, quality });
    return new Promise((resolve, reject) => {
        canvas.toBlob(
            (blob) => blob ? resolve(blob) : reject(new Error("APIMart 参考图编码失败")),
            type,
            quality,
        );
    });
}

async function encodeReferenceImage(image, width, height, quality) {
    const canvas = typeof OffscreenCanvas === "function"
        ? new OffscreenCanvas(width, height)
        : document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("APIMart 参考图压缩失败");
    context.drawImage(image, 0, 0, width, height);
    const webp = await canvasBlob(canvas, "image/webp", quality);
    if (webp.type === "image/webp") return webp;
    const jpeg = await canvasBlob(canvas, "image/jpeg", quality);
    if (jpeg.type === "image/jpeg") return jpeg;
    throw new Error("当前浏览器不支持 APIMart 参考图压缩");
}

export async function prepareApiMartReferenceBlob(blob, options = {}) {
    if (!blob?.size) throw new Error("APIMart 参考图为空");
    if (blob.size <= (options.maxBytes || APIMART_UPLOAD_MAX_BYTES)) return blob;
    const decoded = await (options.decodeImage || decodeReferenceBlob)(blob);
    const encodeImage = options.encodeImage || encodeReferenceImage;
    const targetBytes = options.targetBytes || APIMART_UPLOAD_TARGET_BYTES;
    try {
        let smallest = null;
        for (const scale of REFERENCE_SCALES) {
            const width = Math.max(1, Math.round(decoded.width * scale));
            const height = Math.max(1, Math.round(decoded.height * scale));
            for (const quality of REFERENCE_QUALITIES) {
                throwIfAborted(options.signal);
                const candidate = await encodeImage(decoded.image, width, height, quality);
                if (!smallest || candidate.size < smallest.size) smallest = candidate;
                if (candidate.size <= targetBytes) return candidate;
            }
        }
        if (smallest?.size <= APIMART_UPLOAD_MAX_BYTES) return smallest;
        throw new Error("APIMart 临时参考图压缩后仍超过 20 MB");
    } finally {
        decoded.dispose?.();
    }
}

function uploadFilename(name, blob) {
    const base = String(name || "reference")
        .replace(/\.[^.]+$/, "")
        .normalize("NFKD")
        .replace(/[^A-Za-z0-9_-]+/g, "-")
        .replace(/^-+|-+$/g, "") || "reference";
    const extension = blob.type === "image/webp" ? "webp"
        : blob.type === "image/jpeg" ? "jpg"
            : blob.type === "image/gif" ? "gif" : "png";
    return `${base}.${extension}`;
}

export async function uploadApiMartReferenceBlob(config, originalBlob, options = {}) {
    const apiKey = String(config?.apiKey || "").trim();
    if (!apiKey) throw new Error("请先填写 APIMart API Key");
    if (originalBlob?.type && !IMAGE_TYPES.has(originalBlob.type)) {
        throw new Error("APIMart 仅支持 JPEG、PNG、WebP 和 GIF 图片");
    }
    const requestBlob = await prepareApiMartReferenceBlob(originalBlob, options);
    if (requestBlob.size > APIMART_UPLOAD_MAX_BYTES) throw new Error("APIMart 临时参考图超过 20 MB");

    const form = new FormData();
    form.set("file", requestBlob, uploadFilename(options.filename, requestBlob));
    options.onProgress?.({ stage: "uploading", progress: 2 });
    const timeoutController = new AbortController();
    const timeout = setTimeout(() => timeoutController.abort(), options.timeoutMs || UPLOAD_TIMEOUT_MS);
    const signal = options.signal && typeof AbortSignal.any === "function"
        ? AbortSignal.any([options.signal, timeoutController.signal])
        : options.signal || timeoutController.signal;
    let payload;
    try {
        payload = await fetchJson(`${apiRoot(config?.baseUrl)}/uploads/images`, {
            method: "POST",
            headers: authHeaders(apiKey, ""),
            body: form,
            signal,
        }, options.fetchImpl || fetch);
    } catch (error) {
        if (timeoutController.signal.aborted && !options.signal?.aborted) {
            throw new Error("APIMart 文件上传超时，请检查网络后重试");
        }
        throw error;
    } finally {
        clearTimeout(timeout);
    }
    const url = String(payload?.url || payload?.data?.url || "").trim();
    if (!/^https?:\/\//i.test(url)) throw new Error("APIMart 文件上传成功但没有返回有效 URL");
    options.onProgress?.({ stage: "uploading", progress: 10 });
    return url;
}

function balanceValue(payload, label) {
    const record = payload?.data && typeof payload.data === "object" ? payload.data : payload;
    if (record?.unlimited_quota === true) return -1;
    const raw = record?.remain_balance ?? record?.balance ?? record?.remaining_balance;
    const value = typeof raw === "number" ? raw : Number(raw);
    if (!Number.isFinite(value)) throw new Error(`APIMart ${label}接口没有返回 remain_balance`);
    return value;
}

async function fetchBalance(config, endpoint, label, options) {
    const apiKey = String(config?.apiKey || "").trim();
    if (!apiKey) throw new Error("请先填写 APIMart API Key");
    const payload = await fetchJson(`${apiRoot(config?.baseUrl)}${endpoint}`, {
        method: "GET",
        headers: authHeaders(apiKey),
        cache: "no-store",
        credentials: "omit",
        signal: options.signal,
    }, options.fetchImpl || fetch);
    return balanceValue(payload, label);
}

export function fetchApiMartTokenBalance(config, options = {}) {
    return fetchBalance(config, "/balance", "令牌余额", options);
}

export function fetchApiMartUserBalance(config, options = {}) {
    return fetchBalance(config, "/user/balance", "用户余额", options);
}
