export const APOLLO_BILLING_GROUP_AUTO = "auto";
export const APOLLO_BILLING_GROUP_DEFAULT = "default";
export const APOLLO_BILLING_GROUP_OFFICIAL_MIX = "official-mix";
export const APOLLO_BILLING_GROUP_UNKNOWN = "unknown";

export const APOLLO_BILLING_GROUP_OPTIONS = [
    { value: APOLLO_BILLING_GROUP_AUTO, label: "自动检测（仅明确分组）" },
    { value: APOLLO_BILLING_GROUP_DEFAULT, label: "Default" },
    { value: APOLLO_BILLING_GROUP_OFFICIAL_MIX, label: "Gemini 优质" },
];

export function normalizeApolloBillingGroup(value) {
    const normalized = String(value || "").trim().toLowerCase();
    if (normalized === APOLLO_BILLING_GROUP_DEFAULT) return APOLLO_BILLING_GROUP_DEFAULT;
    if (["official-mix", "gpt-image-2-official-mix", "official_mix"].includes(normalized)) {
        return APOLLO_BILLING_GROUP_OFFICIAL_MIX;
    }
    return APOLLO_BILLING_GROUP_AUTO;
}

export function normalizeApolloDetectedBillingGroup(value) {
    const normalized = normalizeApolloBillingGroup(value);
    return normalized === APOLLO_BILLING_GROUP_AUTO ? APOLLO_BILLING_GROUP_UNKNOWN : normalized;
}

export function apolloBillingGroupFromLabel(label) {
    const normalized = String(label || "").trim().toLowerCase();
    const isDefault = normalized.includes("普通") || normalized.includes("default");
    const isPremium = normalized.includes("优质") || normalized.includes("premium") || normalized.includes("official");
    if (isDefault === isPremium) return APOLLO_BILLING_GROUP_UNKNOWN;
    return isDefault ? APOLLO_BILLING_GROUP_DEFAULT : APOLLO_BILLING_GROUP_OFFICIAL_MIX;
}

export function effectiveApolloBillingGroup(key) {
    const configured = normalizeApolloBillingGroup(key?.billingGroup);
    if (configured !== APOLLO_BILLING_GROUP_AUTO) return configured;
    const namedGroup = apolloBillingGroupFromLabel(key?.label);
    if (namedGroup !== APOLLO_BILLING_GROUP_UNKNOWN) return namedGroup;
    if (["models-capability", "models-default"].includes(key?.billingDetectionSource)) {
        return APOLLO_BILLING_GROUP_UNKNOWN;
    }
    return normalizeApolloDetectedBillingGroup(key?.detectedBillingGroup);
}

export function apolloGptImagePrice(key) {
    const group = effectiveApolloBillingGroup(key);
    if (group === APOLLO_BILLING_GROUP_DEFAULT) return 0.06;
    if (group === APOLLO_BILLING_GROUP_OFFICIAL_MIX) return 0.3;
    return null;
}
export function apolloGptImage25Price(key) {
    const group = effectiveApolloBillingGroup(key);
    if (group === APOLLO_BILLING_GROUP_DEFAULT) return 0.15;
    if (group === APOLLO_BILLING_GROUP_OFFICIAL_MIX) return 0.3;
    return null;
}


export function apolloBillingGroupLabel(key) {
    const group = effectiveApolloBillingGroup(key);
    if (group === APOLLO_BILLING_GROUP_DEFAULT) return "Default";
    if (group === APOLLO_BILLING_GROUP_OFFICIAL_MIX) return "Gemini 优质";
    return "待手动确认";
}

export function apolloBillingDetectionLabel(source) {
    if (source === "quota-group") return "余额接口";
    if (source === "models-group") return "模型接口分组";
    if (source === "label") return "密钥名称";
    if (source === "models-capability" || source === "models-default") return "旧版推断，请重新检测";
    if (source === "manual") return "手动指定";
    return "接口未返回明确分组";
}

function billingGroupFromText(value) {
    const text = String(value || "").trim().toLowerCase();
    if (!text) return APOLLO_BILLING_GROUP_UNKNOWN;
    if (
        text.includes("gpt-image-2-official-mix") ||
        text.includes("official-mix") ||
        text.includes("official_mix") ||
        text.includes("gemini优质") ||
        text.includes("gemini 优质") ||
        text === "premium"
    ) {
        return APOLLO_BILLING_GROUP_OFFICIAL_MIX;
    }
    return text === "default" ? APOLLO_BILLING_GROUP_DEFAULT : APOLLO_BILLING_GROUP_UNKNOWN;
}

function explicitGroupFromPayload(payload) {
    const seen = new Set();
    const groups = [];
    const visit = (value) => {
        if (!value || typeof value !== "object" || seen.has(value)) return;
        seen.add(value);
        if (Array.isArray(value)) {
            value.forEach(visit);
            return;
        }
        for (const [key, child] of Object.entries(value)) {
            const normalizedKey = key.toLowerCase().replace(/[-\s]/g, "_");
            if (
                typeof child === "string" &&
                ["group", "group_name", "billing_group", "billing_group_name", "token_group"].includes(normalizedKey)
            ) {
                groups.push(billingGroupFromText(child));
            }
            visit(child);
        }
    };
    visit(payload);
    if (groups.includes(APOLLO_BILLING_GROUP_OFFICIAL_MIX)) return APOLLO_BILLING_GROUP_OFFICIAL_MIX;
    if (groups.includes(APOLLO_BILLING_GROUP_DEFAULT)) return APOLLO_BILLING_GROUP_DEFAULT;
    return APOLLO_BILLING_GROUP_UNKNOWN;
}

function apolloApiUrl(baseUrl, path) {
    const trimmed = String(baseUrl || "").trim().replace(/\/+$/, "");
    const root = /\/v1$/i.test(trimmed) ? trimmed : `${trimmed}/v1`;
    return `${root}${path}`;
}

async function readJsonResponse(response, label) {
    const payload = await response.json().catch(() => ({}));
    if (response.ok) return payload;
    const message = payload?.error?.message || payload?.message || payload?.msg || `${label}失败 ${response.status}`;
    const error = new Error(message);
    error.status = response.status;
    throw error;
}

export async function detectApolloKeyBillingGroup({ baseUrl, apiKey, fetchImpl = fetch, signal } = {}) {
    const token = String(apiKey || "").trim();
    if (!token) throw new Error("请先填写 Apollo API Key");
    const headers = { Accept: "application/json", Authorization: `Bearer ${token}` };
    const request = (path, label) =>
        fetchImpl(apolloApiUrl(baseUrl, path), {
            method: "GET",
            headers,
            credentials: "omit",
            cache: "no-store",
            signal,
        }).then((response) => readJsonResponse(response, label));

    const [quotaResult, modelsResult] = await Promise.allSettled([
        request("/token/quota", "余额查询"),
        request("/models", "模型权限查询"),
    ]);

    const authFailure = [quotaResult, modelsResult].find(
        (result) => result.status === "rejected" && [401, 403].includes(result.reason?.status),
    );
    if (authFailure) throw authFailure.reason;

    if (quotaResult.status === "rejected" && modelsResult.status === "rejected") {
        throw quotaResult.reason instanceof Error ? quotaResult.reason : new Error("Apollo 密钥检测失败");
    }

    if (quotaResult.status === "fulfilled") {
        const quotaGroup = explicitGroupFromPayload(quotaResult.value);
        if (quotaGroup !== APOLLO_BILLING_GROUP_UNKNOWN) {
            return { group: quotaGroup, source: "quota-group", checkedAt: new Date().toISOString() };
        }
    }

    if (modelsResult.status === "fulfilled") {
        const modelsGroup = explicitGroupFromPayload(modelsResult.value);
        if (modelsGroup !== APOLLO_BILLING_GROUP_UNKNOWN) {
            return { group: modelsGroup, source: "models-group", checkedAt: new Date().toISOString() };
        }
    }

    return { group: APOLLO_BILLING_GROUP_UNKNOWN, source: "unknown", checkedAt: new Date().toISOString() };
}
