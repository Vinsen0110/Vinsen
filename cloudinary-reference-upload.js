export const CLOUDINARY_MAX_BYTES = 10_000_000;
export const CLOUDINARY_TARGET_BYTES = 9_500_000;
export const CLOUDINARY_MAX_PIXELS = 25_000_000;
const UPLOAD_TIMEOUT_MS = 180_000;
const QUALITIES = [0.94, 0.88, 0.8, 0.7, 0.6];

export function usesCloudinaryReferenceHost(config) {
    let hostname = "";
    try { hostname = new URL(config?.baseUrl).hostname; } catch {}
    const supported = config?.provider
        ? ["tudou", "grsai"].includes(config.provider)
        : ["api.ai-tudou.net", "grsaiapi.com", "grsai.dakka.com", "grsai.com"].includes(hostname);
    return supported;
}

export function isCloudinaryImageUrl(value) {
    try {
        const url = new URL(String(value || ""));
        return url.protocol === "https:" && url.hostname === "res.cloudinary.com"
            && !url.username && !url.password && !url.port
            && /^\/[a-zA-Z0-9_-]+\/image\/upload\/.+/.test(url.pathname);
    } catch {
        return false;
    }
}

export function isLegacyImgBbImageUrl(value) {
    try {
        const url = new URL(String(value || ""));
        return url.protocol === "https:" && !url.username && !url.password && !url.port
            && (url.hostname === "ibb.co" || url.hostname.endsWith(".ibb.co"));
    } catch {
        return false;
    }
}

function checkAbort(signal) {
    if (signal?.aborted) throw signal.reason || new DOMException("Aborted", "AbortError");
}

async function decodeImage(blob) {
    if (typeof createImageBitmap === "function") {
        try {
            const image = await createImageBitmap(blob, { imageOrientation: "from-image" });
            return { image, width: image.width, height: image.height, dispose: () => image.close() };
        } catch {}
    }
    const url = URL.createObjectURL(blob);
    try {
        const image = new Image();
        image.src = url;
        await image.decode();
        return {
            image, width: image.naturalWidth, height: image.naturalHeight,
            dispose: () => URL.revokeObjectURL(url),
        };
    } catch (error) {
        URL.revokeObjectURL(url);
        throw error;
    }
}

async function encodeImage(image, width, height, quality) {
    const canvas = typeof OffscreenCanvas === "function"
        ? new OffscreenCanvas(width, height)
        : Object.assign(document.createElement("canvas"), { width, height });
    try {
        const context = canvas.getContext("2d");
        if (!context) throw new Error("无法创建参考图上传副本");
        context.drawImage(image, 0, 0, width, height);
        // WebP retains alpha; do not flatten transparent PNGs onto a JPEG background.
        if (canvas.convertToBlob) return await canvas.convertToBlob({ type: "image/webp", quality });
        return await new Promise((resolve, reject) => canvas.toBlob(
            blob => blob ? resolve(blob) : reject(new Error("参考图副本压缩失败")),
            "image/webp", quality,
        ));
    } finally {
        canvas.width = 0;
        canvas.height = 0;
    }
}

export async function prepareCloudinaryReferenceBlob(original, options = {}) {
    const { signal, onProgress } = options;
    checkAbort(signal);
    if (!original?.size) throw new Error("参考图为空");
    if (!["image/jpeg", "image/png", "image/webp"].includes(original.type.toLowerCase())) {
        throw new Error("Cloudinary 参考图仅支持静态 JPG、PNG、WebP，请先转换格式");
    }
    onProgress?.({ stage: "preparing", progress: 2 });
    const decoded = await (options.decodeImage || decodeImage)(original);
    try {
        checkAbort(signal);
        const { width, height } = decoded;
        if (!(width > 0 && height > 0 && Number.isFinite(width * height))) {
            throw new Error("参考图尺寸无效");
        }
        if (original.size <= CLOUDINARY_MAX_BYTES && width * height <= CLOUDINARY_MAX_PIXELS) {
            return original;
        }
        let scale = Math.min(1, Math.sqrt(CLOUDINARY_MAX_PIXELS / (width * height)));
        const encode = options.encodeImage || encodeImage;
        for (let step = 0; step < 12; step += 1) {
            const targetWidth = Math.max(1, Math.floor(width * scale));
            const targetHeight = Math.max(1, Math.floor(height * scale));
            for (const quality of QUALITIES) {
                checkAbort(signal);
                const blob = await encode(decoded.image, targetWidth, targetHeight, quality);
                checkAbort(signal);
                if (blob?.size > 0 && blob.size <= CLOUDINARY_TARGET_BYTES
                    && ["image/webp", "image/png"].includes(blob.type)) return blob;
            }
            if (Math.max(targetWidth, targetHeight) <= 1024) break;
            scale *= Math.max(0.8, 1024 / Math.max(targetWidth, targetHeight));
        }
        throw new Error("参考图上传副本仍超过限制，请选择较小图片；本地原图未修改");
    } finally {
        decoded.dispose?.();
    }
}

export function cloudinaryCredentials(config) {
    const cloudName = String(config?.cloudinaryCloudName || "").trim();
    const apiKey = String(config?.cloudinaryApiKey || "").trim();
    const apiSecret = String(config?.cloudinaryApiSecret || "").trim();
    if (!cloudName || !apiKey || !apiSecret) {
        throw new Error("请先在 API 设置中填写你自己的 Cloudinary Cloud Name、API Key 和 API Secret");
    }
    if (!/^[a-zA-Z0-9_-]+$/.test(cloudName) || !/^\d+$/.test(apiKey)) {
        throw new Error("Cloudinary Cloud Name 或 API Key 格式不正确");
    }
    return { cloudName, apiKey, apiSecret };
}

export async function signCloudinaryUpload(config) {
    const { cloudName, apiKey, apiSecret } = cloudinaryCredentials(config);
    if (!globalThis.crypto?.subtle || !globalThis.crypto?.randomUUID) {
        throw new Error("当前浏览器无法安全签名，请使用 HTTPS 网站或本地预览");
    }
    const params = {
        overwrite: false,
        public_id: `laowu-reference/${crypto.randomUUID()}`,
        timestamp: Math.floor(Date.now() / 1000),
    };
    const serialized = Object.keys(params).sort().map(key => `${key}=${params[key]}`).join("&");
    const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(serialized + apiSecret));
    const signature = Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, "0")).join("");
    // Only the signature leaves this browser; never send the user's API Secret.
    return { cloudName, apiKey, params, signature };
}

export async function uploadCloudinaryReferenceBlob(config, original, options = {}) {
    const credentials = cloudinaryCredentials(config);
    const uploadConfig = {
        cloudinaryCloudName: credentials.cloudName,
        cloudinaryApiKey: credentials.apiKey,
        cloudinaryApiSecret: credentials.apiSecret,
    };
    checkAbort(options.signal);
    const controller = new AbortController();
    const abort = () => controller.abort(options.signal.reason);
    options.signal?.addEventListener("abort", abort, { once: true });
    const timeout = setTimeout(() => controller.abort(
        new Error("Cloudinary 上传超时，请重试；本地原图未修改"),
    ), options.timeoutMs ?? UPLOAD_TIMEOUT_MS);
    try {
        const signal = controller.signal;
        const blob = await prepareCloudinaryReferenceBlob(original, { ...options, signal });
        const fetchImpl = options.fetchImpl || fetch;
        const signed = await signCloudinaryUpload(uploadConfig);
        checkAbort(signal);
        const params = signed.params;
        const body = new FormData();
        for (const [key, value] of Object.entries(params)) {
            body.set(key, String(value));
        }
        const extension = { "image/png": "png", "image/jpeg": "jpg", "image/webp": "webp" }[blob.type];
        body.set("file", blob, `reference.${extension}`);
        body.set("api_key", String(signed.apiKey));
        body.set("signature", signed.signature);
        options.onProgress?.({ stage: "uploading", progress: 3 });
        // Upload directly into this user's account, with no application signing service.
        const response = await fetchImpl(
            `https://api.cloudinary.com/v1_1/${signed.cloudName}/image/upload`,
            { method: "POST", body, signal, credentials: "omit", redirect: "error" },
        );
        const result = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(result?.error?.message || `Cloudinary 上传失败 (${response.status})`);
        if (!isCloudinaryImageUrl(result.secure_url)
            || new URL(result.secure_url).pathname.split("/")[1] !== signed.cloudName
            || result.public_id !== params.public_id) {
            throw new Error("Cloudinary 未返回有效的参考图链接");
        }
        checkAbort(signal);
        options.onProgress?.({ stage: "uploading", progress: 10 });
        return result.secure_url;
    } catch (error) {
        if (controller.signal.aborted) throw controller.signal.reason;
        throw error;
    } finally {
        clearTimeout(timeout);
        options.signal?.removeEventListener("abort", abort);
    }
}

export async function cloudinaryReferenceSource(config, reference, options = {}) {
    if (!usesCloudinaryReferenceHost(config)) throw new Error("当前站点不使用 Cloudinary 图床");
    checkAbort(options.signal);
    const source = reference.dataUrl || reference.url || "";
    // Preserve old project references. Only local/data/blob sources require a new upload.
    if (isCloudinaryImageUrl(source) || isLegacyImgBbImageUrl(source)) return source;
    let blob = reference.storageKey && options.readStoredBlob
        ? await options.readStoredBlob(reference.storageKey) : null;
    if (!blob) {
        const url = source || await options.readStoredUrl?.(reference.storageKey) || "";
        if (!url) throw new Error("参考图读取失败，请重新选择图片");
        const response = await (options.fetchImpl || fetch)(url, { signal: options.signal });
        if (!response.ok) throw new Error("参考图读取失败，请重新选择图片");
        blob = await response.blob();
    }
    checkAbort(options.signal);
    return uploadCloudinaryReferenceBlob(config, blob, options);
}
