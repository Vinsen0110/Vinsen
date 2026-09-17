import { createHash, randomUUID } from "node:crypto";

const ALLOWED_ORIGINS = new Set([
    "https://vinsen.top", "https://www.vinsen.top", "https://vinsen0110.github.io",
]);
const windows = new Map();

function response(status, payload, origin) {
    return Response.json(payload, {
        status,
        headers: {
            "Cache-Control": "no-store",
            Vary: "Origin",
            ...(ALLOWED_ORIGINS.has(origin) ? {
                "Access-Control-Allow-Origin": origin,
                "Access-Control-Allow-Headers": "Authorization",
                "Access-Control-Allow-Methods": "POST, OPTIONS",
            } : {}),
        },
    });
}

export function createCloudinarySignatureHandler({
    env = process.env, now = Date.now, uuid = randomUUID, rateWindows = windows,
} = {}) {
    return async request => {
        const origin = request.headers.get("origin");
        const error = (status, message) => response(status, { error: { message } }, origin);
        if (origin && !ALLOWED_ORIGINS.has(origin)) return error(403, "不允许的上传来源");
        if (request.method === "OPTIONS") return response(200, {}, origin);
        if (request.method !== "POST") return error(405, "Method Not Allowed");
        const cloudName = String(env.CLOUDINARY_CLOUD_NAME || "");
        const apiKey = String(env.CLOUDINARY_API_KEY || "");
        const secret = String(env.CLOUDINARY_API_SECRET || "");
        const preset = String(env.CLOUDINARY_UPLOAD_PRESET || "");
        const tokenHashes = String(env.CLOUDINARY_UPLOAD_TOKEN_HASHES || "")
            .split(",").map(value => value.trim()).filter(value => /^[a-f0-9]{64}$/.test(value));
        if (!/^[a-zA-Z0-9_-]+$/.test(cloudName) || !/^\d+$/.test(apiKey)
            || !secret || !/^[a-zA-Z0-9_-]+$/.test(preset) || !tokenHashes.length) {
            return error(503, "Cloudinary 服务端签名未配置，请联系管理员");
        }
        const authorization = request.headers.get("authorization") || "";
        if (!/^Bearer [^\s]{32,512}$/.test(authorization)) return error(401, "上传服务访问码无效");
        const hash = createHash("sha256").update(authorization.slice(7)).digest("hex");
        if (!tokenHashes.includes(hash)) return error(401, "上传服务访问码无效或已撤销");
        // This per-instance burst limit supplements authorization, not an account-wide quota.
        const time = now();
        for (const [key, entry] of rateWindows) if (time >= entry.endsAt) rateWindows.delete(key);
        const current = rateWindows.get(hash) || { endsAt: time + 60_000, count: 0 };
        if (current.count >= 20) return error(429, "参考图上传过于频繁，请稍后重试");
        current.count += 1;
        rateWindows.set(hash, current);
        // No caller-supplied transformations, folders, remote URLs, or overwrite controls are signed.
        const params = {
            overwrite: false,
            public_id: `laowu-reference/${uuid()}`,
            tags: "laowu-reference",
            timestamp: Math.floor(time / 1000),
            upload_preset: preset,
        };
        const serialized = Object.keys(params).sort().map(key => `${key}=${params[key]}`).join("&");
        const signature = createHash("sha256").update(serialized + secret).digest("hex");
        return response(200, { cloudName, apiKey, params, signature }, origin);
    };
}

export default { fetch: createCloudinarySignatureHandler() };
