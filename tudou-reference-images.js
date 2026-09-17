import { isCloudinaryImageUrl, isLegacyImgBbImageUrl } from "./cloudinary-reference-upload.js";

const MAX_REFERENCE_BYTES = 14 * 1024 * 1024;
const TUDOU_GEMINI_IMAGE_PATH = /^\/v1beta\/models\/[^/]+:(?:generateContent|streamGenerateContent)$/;

export function isImgBbImageUrl(value) {
    return isLegacyImgBbImageUrl(value);
}

export function isHostedReferenceImageUrl(value) {
    return isImgBbImageUrl(value) || isCloudinaryImageUrl(value);
}

export function isTudouGeminiImageTarget(target) {
    const url = target instanceof URL ? target : new URL(String(target));
    return TUDOU_GEMINI_IMAGE_PATH.test(url.pathname);
}

function imageMimeType(response) {
    const type = String(response.headers.get("content-type") || "")
        .split(";", 1)[0]
        .trim()
        .toLowerCase();
    return type.startsWith("image/") ? type : "";
}

async function inlineHostedImage(fileUri, signal, fetchImpl) {
    let target = fileUri;
    let response;
    for (let redirects = 0; redirects <= 3; redirects += 1) {
        if (!isHostedReferenceImageUrl(target)) throw new Error("Reference redirected to an unsupported host");
        response = await fetchImpl(target, {
            headers: { Accept: "image/*", "Accept-Encoding": "identity" },
            redirect: "manual", signal,
        });
        if (![301, 302, 303, 307, 308].includes(response.status)) break;
        const location = response.headers.get("location");
        await response.body?.cancel();
        if (!location || redirects === 3) throw new Error("Too many reference image redirects");
        target = new URL(location, target).href;
    }
    if (!response.ok) throw new Error(`Reference download failed (${response.status})`);
    if (!isHostedReferenceImageUrl(response.url || target)) {
        throw new Error("Reference redirected to an unsupported host");
    }

    const mimeType = imageMimeType(response);
    if (!mimeType) {
        await response.body?.cancel();
        throw new Error("Reference did not return an image");
    }
    const declaredBytes = Number(response.headers.get("content-length")) || 0;
    if (declaredBytes > MAX_REFERENCE_BYTES) {
        await response.body?.cancel();
        throw new Error("Reference exceeds the 14 MB limit");
    }

    const chunks = [];
    let total = 0;
    const reader = response.body?.getReader();
    if (!reader) throw new Error("Reference download returned no image");
    try {
        for (;;) {
            const { done, value } = await reader.read();
            if (done) break;
            total += value.byteLength;
            if (total > MAX_REFERENCE_BYTES) {
                await reader.cancel();
                throw new Error("Reference exceeds the 14 MB limit");
            }
            chunks.push(value);
        }
    } finally {
        reader.releaseLock();
    }
    if (!total) throw new Error("Reference download returned an empty image");
    return { mimeType, data: Buffer.concat(chunks, total).toString("base64") };
}

export async function inlineTudouGeminiReferences(target, payload, options = {}) {
    if (!isTudouGeminiImageTarget(target) || !payload || typeof payload !== "object") {
        return { payload, converted: 0 };
    }

    const fetchImpl = options.fetchImpl || fetch;
    let converted = 0;
    for (const content of Array.isArray(payload.contents) ? payload.contents : []) {
        for (const part of Array.isArray(content?.parts) ? content.parts : []) {
            const fileData = part?.fileData || part?.file_data;
            const fileUri = fileData?.fileUri || fileData?.file_uri;
            if (!isHostedReferenceImageUrl(fileUri)) continue;

            part.inlineData = await inlineHostedImage(fileUri, options.signal, fetchImpl);
            delete part.fileData;
            delete part.file_data;
            converted += 1;
        }
    }
    return { payload, converted };
}
