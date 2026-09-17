import { createReadStream, promises as fs } from "node:fs";
import { createServer } from "node:http";
import { extname, resolve, sep } from "node:path";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { fetchApiMartResultImage } from "./apimart-result-image.js";
import { inlineTudouGeminiReferences, isTudouGeminiImageTarget } from "./tudou-reference-images.js";

const root = resolve(process.cwd());
const host = process.env.HOST || "127.0.0.1";
const port = Number(process.env.PORT || 4173);
const tudouOrigin = "https://api.ai-tudou.net";
const allowedTudouPath = /^\/(?:images|tasks|v1|v1beta)(?:\/|$)/;

const mimeTypes = {
    ".css": "text/css; charset=utf-8",
    ".html": "text/html; charset=utf-8",
    ".ico": "image/x-icon",
    ".jpeg": "image/jpeg",
    ".jpg": "image/jpeg",
    ".js": "text/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
    ".webp": "image/webp",
};

function sendJson(response, status, message) {
    response.writeHead(status, {
        "Cache-Control": "no-store",
        "Content-Type": "application/json; charset=utf-8",
    });
    response.end(JSON.stringify({ error: { message } }));
}

function errorDetails(error) {
    if (!(error instanceof Error)) return String(error);
    const cause = error.cause instanceof Error
        ? `; cause=${error.cause.name}: ${error.cause.message}${error.cause.code ? ` (${error.cause.code})` : ""}`
        : "";
    return `${error.name}: ${error.message}${cause}`;
}

function requestHeaders(request) {
    const headers = new Headers({
        Accept: request.headers.accept || "application/json",
        "Accept-Encoding": "identity",
        Authorization: request.headers.authorization,
    });
    if (request.headers["content-type"]) {
        headers.set("Content-Type", request.headers["content-type"]);
    }
    return headers;
}

async function readRequestBody(request, maxBytes = 8 * 1024 * 1024) {
    const chunks = [];
    let bytes = 0;
    for await (const chunk of request) {
        const value = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
        bytes += value.byteLength;
        if (bytes > maxBytes) throw new Error("Tudou proxy request body is too large");
        chunks.push(value);
    }
    return Buffer.concat(chunks);
}

async function pipeFetchResponse(upstream, response, extraHeaders = {}) {
    response.statusCode = upstream.status;
    response.statusMessage = upstream.statusText;
    response.setHeader("Cache-Control", "no-store");
    for (const name of ["content-type", "retry-after", "x-oneapi-request-id", "x-request-id"]) {
        const value = upstream.headers.get(name);
        if (value) response.setHeader(name, value);
    }
    for (const [name, value] of Object.entries(extraHeaders)) response.setHeader(name, value);
    if (!upstream.body) {
        response.end();
        return;
    }
    await pipeline(Readable.fromWeb(upstream.body), response);
}

function sseData(value) {
    return `data: ${JSON.stringify(value)}\n\n`;
}

async function proxyTudouGeminiStream(request, response, target, headers, payload, controller) {
    let closed = false;
    const heartbeat = setInterval(() => {
        if (!closed && !response.writableEnded) response.write(": tudou-proxy\n\n");
    }, 5_000);
    response.writeHead(200, {
        "Cache-Control": "no-store",
        "Content-Type": "text/event-stream; charset=utf-8",
        "X-Accel-Buffering": "no",
        "X-Tudou-Proxy": "local-node-stream",
        "X-Tudou-References-Inlined": "pending",
    });
    response.write(": tudou-proxy\n\n");

    try {
        const normalized = await inlineTudouGeminiReferences(target, payload, {
            signal: controller.signal,
        });
        const upstream = await fetch(target, {
            method: "POST",
            headers,
            body: JSON.stringify(normalized.payload),
            redirect: "manual",
            signal: controller.signal,
        });
        // Once upstream bytes start flowing, comments must never split an SSE frame.
        clearInterval(heartbeat);
        if (!upstream.ok) {
            const raw = await upstream.text();
            let message = raw || `Tudou request failed (${upstream.status})`;
            try {
                const parsed = JSON.parse(raw);
                message = parsed?.error?.message || parsed?.message || message;
            } catch {
                // Keep the upstream text when it is not JSON.
            }
            response.write(sseData({ error: { message } }));
            return;
        }

        if (!String(upstream.headers.get("content-type") || "").includes("text/event-stream")) {
            const raw = await upstream.text();
            try {
                response.write(sseData(JSON.parse(raw)));
            } catch {
                response.write(sseData({ error: { message: raw || "Tudou returned an empty response" } }));
            }
            return;
        }

        if (!upstream.body) {
            response.write(sseData({ error: { message: "Tudou stream returned no body" } }));
            return;
        }
        const reader = upstream.body.getReader();
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            if (!closed && !response.writableEnded && value) response.write(Buffer.from(value));
        }
    } catch (error) {
        if (!controller.signal.aborted && !closed && !response.writableEnded) {
            console.error("Tudou streaming proxy request failed:", errorDetails(error));
            response.write(sseData({ error: { message: errorDetails(error) } }));
        }
    } finally {
        closed = true;
        clearInterval(heartbeat);
        if (!response.writableEnded) response.end();
    }
}

async function proxyTudouApi(request, response, requestUrl) {
    const method = request.method?.toUpperCase() || "GET";
    if (!["GET", "HEAD", "OPTIONS", "POST"].includes(method)) {
        sendJson(response, 405, "Method Not Allowed");
        return;
    }
    if (method === "OPTIONS") {
        response.writeHead(204, { Allow: "GET, HEAD, OPTIONS, POST" });
        response.end();
        return;
    }

    const path = requestUrl.pathname.slice("/api/tudou".length) || "/";
    if (!allowedTudouPath.test(path) || path.split(/[\\/]/).includes("..")) {
        sendJson(response, 404, "Unsupported Tudou API path");
        return;
    }
    if (!request.headers.authorization) {
        sendJson(response, 401, "Missing Authorization header");
        return;
    }

    const target = new URL(path, tudouOrigin);
    requestUrl.searchParams.forEach((value, key) => target.searchParams.append(key, value));
    const controller = new AbortController();
    request.on("aborted", () => controller.abort());
    response.on("close", () => {
        if (!response.writableEnded) controller.abort();
    });

    try {
        const headers = requestHeaders(request);
        const isGeminiImageRequest = method === "POST"
            && isTudouGeminiImageTarget(target)
            && String(request.headers["content-type"] || "").includes("application/json");
        if (isGeminiImageRequest) {
            const payload = JSON.parse((await readRequestBody(request)).toString("utf8"));
            await proxyTudouGeminiStream(request, response, target, headers, payload, controller);
            return;
        }

        let body;
        if (method !== "GET" && method !== "HEAD") {
            body = request;
        }
        const upstream = await fetch(target, {
            method,
            headers,
            body,
            duplex: body === request ? "half" : undefined,
            redirect: "manual",
            signal: controller.signal,
        });
        await pipeFetchResponse(upstream, response, {
            "X-Tudou-Proxy": "local-node-stream",
            "X-Tudou-References-Inlined": "0",
        });
    } catch (error) {
        if (controller.signal.aborted) return;
        console.error("Tudou API proxy failed:", errorDetails(error));
        if (!response.headersSent) sendJson(response, 502, "Tudou API connection failed");
        else response.destroy(error);
    }
}

async function proxyTudouImage(response, requestUrl) {
    try {
        const target = new URL(requestUrl.searchParams.get("url") || "");
        const allowedHost = target.hostname === "ai-tudou.net" || target.hostname.endsWith(".ai-tudou.net");
        if (target.protocol !== "https:" || !allowedHost || target.hostname === "api.ai-tudou.net") {
            sendJson(response, 400, "Unsupported Tudou image URL");
            return;
        }
        const upstream = await fetch(target, {
            headers: { Accept: "image/*", "Accept-Encoding": "identity" },
            redirect: "follow",
        });
        await pipeFetchResponse(upstream, response, { "X-Tudou-Image-Proxy": "local-node-stream" });
    } catch (error) {
        console.error("Tudou image proxy failed:", errorDetails(error));
        if (!response.headersSent) sendJson(response, 502, "Tudou image connection failed");
        else response.destroy(error);
    }
}

async function proxyApiMartImage(response, requestUrl) {
    try {
        const upstream = await fetchApiMartResultImage(requestUrl.searchParams.get("url"));
        await pipeFetchResponse(upstream, response, { "X-APIMart-Image-Proxy": "local-node-stream" });
    } catch (error) {
        console.error("APIMart image proxy failed:", errorDetails(error));
        if (!response.headersSent) sendJson(response, 502, "APIMart image connection failed");
        else response.destroy(error);
    }
}

async function serveStatic(request, response, requestUrl) {
    let pathname;
    try {
        pathname = decodeURIComponent(requestUrl.pathname);
    } catch {
        sendJson(response, 400, "Invalid URL");
        return;
    }

    const relativePath = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
    const isIconAsset = pathname === "/icons" || pathname.startsWith("/icons/");
    let filePath = resolve(root, relativePath);
    if (filePath !== root && !filePath.startsWith(`${root}${sep}`)) {
        sendJson(response, 403, "Forbidden");
        return;
    }

    try {
        const info = await fs.stat(filePath);
        if (!info.isFile()) throw new Error("Not a file");
    } catch {
        filePath = resolve(root, "index.html");
    }

    const info = await fs.stat(filePath);
    response.writeHead(200, {
        "Cache-Control": isIconAsset
            ? "public, max-age=31536000, immutable"
            : "no-store",
        "Content-Length": info.size,
        "Content-Type": mimeTypes[extname(filePath).toLowerCase()] || "application/octet-stream",
    });
    if (request.method === "HEAD") response.end();
    else await pipeline(createReadStream(filePath), response);
}

const server = createServer(async (request, response) => {
    const requestUrl = new URL(request.url || "/", `http://${request.headers.host || `${host}:${port}`}`);
    try {
        if (requestUrl.pathname.startsWith("/api/tudou/")) {
            await proxyTudouApi(request, response, requestUrl);
        } else if (requestUrl.pathname === "/api/tudou-image") {
            await proxyTudouImage(response, requestUrl);
        } else if (requestUrl.pathname === "/api/apimart-image") {
            await proxyApiMartImage(response, requestUrl);
        } else {
            await serveStatic(request, response, requestUrl);
        }
    } catch (error) {
        console.error("Local preview request failed:", errorDetails(error));
        if (!response.headersSent) sendJson(response, 500, "Local preview request failed");
        else response.destroy(error);
    }
});

server.listen(port, host, () => {
    console.log(`Local preview with Tudou streaming proxy: http://${host}:${port}/`);
});
