export const GRSAI_ORIGIN = "https://grsaiapi.com";
export const GRSAI_SITE_ID = "grsai";
export const GRSAI_SITE_NAME = "Grsai";
export const GRSAI_SITE_MODELS = ["nano-banana-pro"];
export const GRSAI_IMAGE_MODELS = GRSAI_SITE_MODELS;
export const GRSAI_TEXT_MODELS = [];

const POLL_INTERVAL_MS = 2500;
const TASK_TIMEOUT_MS = 5 * 60 * 1000;

const VIP_RATIO_SIZES = {
    "1:1": { "1k": "1024x1024", "2k": "2048x2048", "4k": "2880x2880" },
    "16:9": { "1k": "1280x720", "2k": "2048x1152", "4k": "3840x2160" },
    "9:16": { "1k": "720x1280", "2k": "1152x2048", "4k": "2160x3840" },
    "4:3": { "1k": "1152x864", "2k": "2304x1728", "4k": "3264x2448" },
    "3:4": { "1k": "864x1152", "2k": "1728x2304", "4k": "2448x3264" },
    "3:2": { "1k": "1536x1024", "2k": "2048x1360", "4k": "3504x2336" },
    "2:3": { "1k": "1024x1536", "2k": "1360x2048", "4k": "2336x3504" },
    "5:4": { "1k": "1120x896", "2k": "2240x1792", "4k": "3200x2560" },
    "4:5": { "1k": "896x1120", "2k": "1792x2240", "4k": "2560x3200" },
    "21:9": { "1k": "1456x624", "2k": "2912x1248", "4k": "3840x1648" },
    "9:21": { "1k": "624x1456", "2k": "1248x2912", "4k": "1648x3840" },
    "1:3": { "1k": "688x2048", "2k": "1280x3840", "4k": "1280x3840" },
    "3:1": { "1k": "2048x688", "2k": "3840x1280", "4k": "3840x1280" },
    "2:1": { "1k": "1536x768", "2k": "3072x1536", "4k": "3840x1920" },
    "1:2": { "1k": "768x1536", "2k": "1536x3072", "4k": "1920x3840" },
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

function siteOrigin(config) {
    return String(config?.baseUrl || GRSAI_ORIGIN).trim().replace(/\/+$/, "");
}

export function isGrsaiSite(config) {
    return config?.provider === "grsai"
        || /grsai(?:api|\.dakka)\.com|grsai\.com/i.test(String(config?.baseUrl || ""));
}

export function grsaiResolution(config) {
    const value = String(config?.quality || "1k").trim().toLowerCase();
    if (["1k", "2k", "4k"].includes(value)) return value;
    if (["low", "standard", "auto"].includes(value)) return "1k";
    if (["medium", "hd"].includes(value)) return "2k";
    if (value === "high") return "4k";
    return "1k";
}

export function grsaiImageRequestSpec(config, prompt, imageUrls = []) {
    const model = String(config?.model || config?.imageModel || "nano-banana-pro").trim()
        .split("::").at(-1);
    const resolution = grsaiResolution(config);
    const ratio = String(config?.size || "auto").trim();
    const isVip = model === "gpt-image-2-vip";
    const aspectRatio = isVip ? (VIP_RATIO_SIZES[ratio]?.[resolution] || "auto") : ratio;
    const body = {
        model: isVip ? "gpt-image-2-vip" : "nano-banana-pro",
        prompt: String(prompt || "").trim(),
        images: Array.isArray(imageUrls) ? imageUrls : [],
        aspectRatio,
        replyType: "async",
        ...(isVip ? {} : { imageSize: resolution.toUpperCase() }),
    };
    return { endpoint: "/v1/api/generate", body };
}

function jsonHeaders(apiKey) {
    return {
        Authorization: `Bearer ${String(apiKey || "").trim()}`,
        "Content-Type": "application/json",
        Accept: "application/json",
    };
}

async function fetchJson(url, init, fetchImpl) {
    const response = await fetchImpl(url, init);
    const text = await response.text();
    let payload = {};
    try {
        payload = text ? JSON.parse(text) : {};
    } catch {
        payload = { error: text || `HTTP ${response.status}` };
    }
    if (!response.ok) {
        const message = payload?.message || payload?.error || payload?.msg || `Grsai 请求失败（HTTP ${response.status}）`;
        throw new Error(String(message));
    }
    if (payload?.success === false || payload?.code && payload.code !== 200 && payload.code !== 0) {
        throw new Error(String(payload?.message || payload?.msg || "Grsai 请求失败"));
    }
    return payload;
}

function record(payload) {
    return payload?.data && typeof payload.data === "object" && !Array.isArray(payload.data)
        ? payload.data
        : payload || {};
}

export function parseGrsaiTask(payload) {
    const item = record(payload);
    const id = String(item.id || item.taskId || item.task_id || "").trim();
    const status = String(item.status || "").trim().toLowerCase();
    if (status === "succeeded" || status === "success" || status === "completed") {
        const urls = (Array.isArray(item.results) ? item.results : [])
            .map((result) => String(result?.url || result?.image_url || "").trim())
            .filter((url) => /^https?:\/\//i.test(url));
        return urls.length
            ? { status: "completed", taskId: id, urls }
            : { status: "failed", taskId: id, error: "Grsai 任务成功但没有返回图片" };
    }
    if (status === "failed" || status === "violation") {
        return { status: "failed", taskId: id, error: String(item.error || item.message || `Grsai 任务${status}`) };
    }
    return { status: "pending", taskId: id, progress: Number.isFinite(Number(item.progress)) ? Number(item.progress) : undefined };
}

export async function runGrsaiImageGeneration(config, prompt, imageUrls = [], options = {}) {
    const apiKey = String(config?.apiKey || "").trim();
    if (!apiKey) throw new Error("请先填写 Grsai API Key");
    if (!String(prompt || "").trim()) throw new Error("请输入提示词");
    const fetchImpl = options.fetchImpl || fetch;
    const sleep = options.sleep || wait;
    const now = options.now || Date.now;
    const submitted = await fetchJson(`${siteOrigin(config)}${grsaiImageRequestSpec(config, prompt, imageUrls).endpoint}`, {
        method: "POST",
        headers: jsonHeaders(apiKey),
        body: JSON.stringify(grsaiImageRequestSpec(config, prompt, imageUrls).body),
        signal: options.signal,
    }, fetchImpl);
    const first = parseGrsaiTask(submitted);
    if (!first.taskId) throw new Error("Grsai 接口没有返回任务 ID");
    options.onProgress?.({ progress: 12, stage: "generating", taskId: first.taskId });
    const startedAt = now();
    for (;;) {
        throwIfAborted(options.signal);
        if (now() - startedAt > TASK_TIMEOUT_MS) throw new Error("Grsai 图片生成超时，请稍后重试");
        const result = parseGrsaiTask(await fetchJson(
            `${siteOrigin(config)}/v1/api/result?id=${encodeURIComponent(first.taskId)}`,
            { method: "GET", headers: jsonHeaders(apiKey), signal: options.signal },
            fetchImpl,
        ));
        if (result.status === "completed") {
            options.onProgress?.({ progress: 100, stage: "completed", taskId: first.taskId });
            return result.urls;
        }
        if (result.status === "failed") throw new Error(result.error);
        if (Number.isFinite(result.progress)) options.onProgress?.({ progress: Math.max(12, Math.min(99, result.progress)), stage: "generating", taskId: first.taskId });
        await sleep(POLL_INTERVAL_MS, options.signal);
    }
}

function creditsFrom(payload) {
    const item = record(payload);
    const value = item.credits ?? item.balance ?? item.quota;
    const number = Number(value);
    return Number.isFinite(number) ? number : null;
}

export async function fetchGrsaiApiKeyCredits(config, options = {}) {
    const apiKey = String(config?.apiKey || "").trim();
    if (!apiKey) throw new Error("请先填写 Grsai API Key");
    const payload = await fetchJson(`${siteOrigin(config)}/client/openapi/getAPIKeyCredits`, {
        method: "POST", headers: jsonHeaders(apiKey), body: JSON.stringify({ apiKey }), signal: options.signal,
    }, options.fetchImpl || fetch);
    const credits = creditsFrom(payload);
    if (credits == null) throw new Error("Grsai API Key 余额接口没有返回 credits");
    return credits;
}

export async function fetchGrsaiAccountCredits(config, token, options = {}) {
    const accountToken = String(token || "").trim();
    if (!accountToken) throw new Error("请填写 Grsai 账户 Token");
    const payload = await fetchJson(`${siteOrigin(config)}/client/openapi/getCredits`, {
        method: "POST", headers: jsonHeaders(config?.apiKey), body: JSON.stringify({ token: accountToken }), signal: options.signal,
    }, options.fetchImpl || fetch);
    const credits = creditsFrom(payload);
    if (credits == null) throw new Error("Grsai 账户余额接口没有返回 credits");
    return credits;
}
