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

function signatureEndpoint() {
    return (typeof location !== "undefined" && location.hostname === "vinsen0110.github.io"
        ? "https://www.vinsen.top" : "") + "/api/cloudinary-signature";
}

export async function uploadCloudinaryReferenceBlob(config, original, options = {}) {
    const token = String(config?.cloudinaryUploadToken || "").trim();
    if (!token) throw new Error("请先配置上传服务，并在 API 设置中填写上传服务访问码");
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
        const signedResponse = await fetchImpl(signatureEndpoint(), {
            method: "POST", headers: { Authorization: `Bearer ${token}` },
            credentials: "omit", cache: "no-store", redirect: "error", signal,
        });
        const signed = await signedResponse.json().catch(() => ({}));
        if (!signedResponse.ok) throw new Error(
            signed?.error?.message || "Cloudinary 签名服务尚未配置或不可用",
        );
        const params = signed.params;
        if (!/^[a-zA-Z0-9_-]+$/.test(signed.cloudName || "")
            || !/^\d+$/.test(String(signed.apiKey || ""))
            || !/^[a-f0-9]{40,64}$/.test(signed.signature || "")
            || !params || typeof params !== "object" || Array.isArray(params)
            || !params.upload_preset || !params.public_id || String(params.overwrite) !== "false"
            || Math.abs(Date.now() / 1000 - Number(params.timestamp)) > 300
            || !Number.isFinite(Number(params.timestamp))) {
            throw new Error("Cloudinary 签名响应无效");
        }
        const allowed = new Set(["timestamp", "upload_preset", "public_id", "overwrite", "tags"]);
        const body = new FormData();
        for (const [key, value] of Object.entries(params)) {
            if (!allowed.has(key) || !["string", "number", "boolean"].includes(typeof value)) {
                throw new Error("Cloudinary 签名包含不支持的参数");
            }
            body.set(key, String(value));
        }
        const extension = { "image/png": "png", "image/jpeg": "jpg", "image/webp": "webp" }[blob.type];
        body.set("file", blob, `reference.${extension}`);
        body.set("api_key", String(signed.apiKey));
        body.set("signature", signed.signature);
        options.onProgress?.({ stage: "uploading", progress: 3 });
        // Never forward the application upload credential or provider API key to the image host.
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
