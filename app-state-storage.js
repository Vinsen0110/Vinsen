export const APP_STATE_KEYS = [
    "infinite-canvas:ai_config_store",
    "infinite-canvas:theme_store",
    "infinite-canvas:asset_store",
    "infinite-canvas:canvas_store",
];
const NATIVE_PREFERENCE_KEYS = [
    "canvas-agent-panel-width",
    "canvas-agent-url",
    "canvas-agent-token",
    "canvas-image-prompt-presets-v1",
    "canvas-image-quick-tools-v7",
];
const MIGRATION_KEY = "migration-v1";

function storageError(name, message) {
    return Object.assign(new Error(message), { name });
}

function openDatabase(indexedDB, name, create, timeoutMs) {
    return new Promise((resolve, reject) => {
        let finished = false;
        const request = indexedDB.open(name);
        const timer = setTimeout(() => {
            finished = true;
            reject(storageError("StorageOpenError", "Browser storage is busy or unavailable"));
        }, timeoutMs);
        request.onupgradeneeded = () => {
            if (!create) {
                request.transaction.abort();
                return;
            }
            const db = request.result;
            for (const store of ["records", "metadata", "backups"]) {
                if (!db.objectStoreNames.contains(store)) db.createObjectStore(store);
            }
        };
        request.onsuccess = () => {
            clearTimeout(timer);
            if (finished) return request.result.close();
            finished = true;
            resolve(request.result);
        };
        request.onerror = () => {
            clearTimeout(timer);
            if (finished) return;
            finished = true;
            if (!create && request.error?.name === "AbortError") resolve(null);
            else reject(request.error);
        };
    });
}

function transactionResult(db, stores, mode, run) {
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(stores, mode);
        let result;
        let failure;
        transaction.oncomplete = () => resolve(result);
        transaction.onabort = () => reject(failure || transaction.error || new Error("Storage transaction aborted"));
        transaction.onerror = () => {};
        try {
            run(transaction, value => { result = value; }, error => {
                failure = error;
                transaction.abort();
            });
        } catch (error) {
            failure = error;
            transaction.abort();
        }
    });
}

async function readKeys(db, store, keys) {
    if (!db?.objectStoreNames.contains(store)) return {};
    return transactionResult(db, [store], "readonly", (transaction, done) => {
        const values = {};
        done(values);
        for (const key of keys) {
            const request = transaction.objectStore(store).get(key);
            request.onsuccess = () => { values[key] = request.result; };
        }
    });
}

function validState(key, value) {
    if (typeof value !== "string") return false;
    try {
        const state = JSON.parse(value)?.state;
        if (!state || typeof state !== "object" || Array.isArray(state)) return false;
        if (key.endsWith(":ai_config_store")) return !!state.config && typeof state.config === "object" && !Array.isArray(state.config);
        if (key.endsWith(":theme_store")) return state.theme === "light" || state.theme === "dark";
        if (key.endsWith(":canvas_store")) return Array.isArray(state.projects) && state.projects.every(project =>
            project && typeof project === "object" && typeof project.id === "string" &&
            typeof project.title === "string" && Array.isArray(project.nodes) && Array.isArray(project.connections));
        if (key.endsWith(":asset_store")) return Array.isArray(state.assets) && state.assets.every(asset =>
            asset && typeof asset === "object" && typeof asset.id === "string" && asset.data && typeof asset.data === "object");
    } catch {}
    return false;
}

function validPreference(key, value) {
    if (typeof value !== "string") return false;
    if (key === "canvas-agent-panel-width") return Number.isFinite(Number(value)) && Number(value) > 0;
    if (key === "canvas-agent-token" || key === "canvas-agent-url") return true;
    try {
        const parsed = JSON.parse(value);
        return key === "canvas-image-prompt-presets-v1"
            ? Array.isArray(parsed)
            : !!parsed && typeof parsed === "object" && !Array.isArray(parsed);
    } catch { return false; }
}

export function createAppStateStorage({
    indexedDB,
    localStorage,
    databaseName = "laowu-app-state",
    timeoutMs = 8000,
    onError = () => {},
} = {}) {
    const revisions = new Map();
    const failures = new Map();
    let db;
    let tail = Promise.resolve();
    let closed = false;

    function report(error, key) {
        failures.set(key, error);
        onError(error);
    }

    const ready = (async () => {
        // Access may itself throw when the browser denies storage permissions.
        if (indexedDB === undefined) indexedDB = globalThis.indexedDB;
        if (localStorage === undefined) localStorage = globalThis.localStorage;
        if (!indexedDB) throw storageError("StorageOpenError", "IndexedDB is unavailable");
        db = await openDatabase(indexedDB, databaseName, true, timeoutMs);
        db.onversionchange = () => { closed = true; db.close(); };
        const metadata = await readKeys(db, "metadata", [MIGRATION_KEY]);
        if (!metadata[MIGRATION_KEY]) {
            const allKeys = [...APP_STATE_KEYS, ...NATIVE_PREFERENCE_KEYS];
            const native = {};
            // Abort migration if native storage cannot be read; defaults must not replace unread data.
            for (const key of allKeys) native[key] = localStorage.getItem(key);
            let previousDb;
            let mirrorDb;
            let previous;
            let mirror;
            try {
                previousDb = await openDatabase(indexedDB, "infinite-canvas", false, timeoutMs);
                mirrorDb = await openDatabase(indexedDB, "vinsen-storage", false, timeoutMs);
                previous = await readKeys(previousDb, "app_state", APP_STATE_KEYS);
                mirror = await readKeys(mirrorDb, "kv", allKeys);
            } finally {
                previousDb?.close();
                mirrorDb?.close();
            }
            const selected = {};
            for (const key of APP_STATE_KEYS) {
                const candidates = key.endsWith(":canvas_store") || key.endsWith(":asset_store")
                    ? [previous[key], native[key], mirror[key]]
                    : [native[key], mirror[key]];
                const value = candidates.find(candidate => validState(key, candidate));
                if (value === undefined && candidates.some(candidate => candidate != null)) {
                    throw storageError("StorageMigrationError", "Existing saved data needs recovery; it has not been overwritten");
                }
                selected[key] = value ?? null;
            }
            await transactionResult(db, ["records", "metadata", "backups"], "readwrite", (transaction, done) => {
                const check = transaction.objectStore("metadata").get(MIGRATION_KEY);
                check.onsuccess = () => {
                    // Another tab may have completed the same migration while we read legacy data.
                    if (check.result) return done(false);
                    transaction.objectStore("backups").put({
                        createdAt: new Date().toISOString(), native, previous, mirror,
                    }, "before-migration-v1");
                    for (const key of APP_STATE_KEYS) {
                        transaction.objectStore("records").put({ value: selected[key], revision: 0 }, key);
                    }
                    transaction.objectStore("metadata").put({ version: 1 }, MIGRATION_KEY);
                    done(true);
                };
            });
        }
        // These small preferences keep native, synchronous Storage semantics.
        // Recover only once, never resurrect a later deliberate deletion.
        const preferences = await readKeys(db, "metadata", ["preferences-v1"]);
        if (!preferences["preferences-v1"]) {
            const backup = (await readKeys(db, "backups", ["before-migration-v1"]))["before-migration-v1"];
            let restored = true;
            for (const key of NATIVE_PREFERENCE_KEYS) {
                const value = backup?.mirror?.[key];
                try {
                    if (localStorage.getItem(key) === null && validPreference(key, value)) localStorage.setItem(key, value);
                } catch (error) {
                    restored = false;
                    onError(error);
                }
            }
            if (restored) await transactionResult(db, ["metadata"], "readwrite", transaction => {
                transaction.objectStore("metadata").put(true, "preferences-v1");
            });
        }
        const records = await readKeys(db, "records", APP_STATE_KEYS);
        for (const key of APP_STATE_KEYS) revisions.set(key, records[key]?.revision ?? 0);
    })();
    ready.catch(error => report(error, "startup"));

    function assertKey(key) {
        if (!APP_STATE_KEYS.includes(key)) throw new Error("Unsupported application storage key");
        if (closed) throw storageError("StorageOpenError", "Storage changed in another page; reload before saving");
    }

    function enqueue(key, operation) {
        const job = tail.then(() => ready).then(operation);
        tail = job.catch(() => {});
        job.catch(error => report(error, key));
        return job;
    }

    async function getItem(key) {
        assertKey(key);
        await tail;
        await ready;
        const record = (await readKeys(db, "records", [key]))[key];
        revisions.set(key, record?.revision ?? 0);
        return record?.value ?? null;
    }

    function write(key, value) {
        assertKey(key);
        const job = enqueue(key, () => transactionResult(db, ["records"], "readwrite", (transaction, done, fail) => {
            const store = transaction.objectStore("records");
            const request = store.get(key);
            request.onsuccess = () => {
                const current = request.result || { value: null, revision: 0 };
                if (current.value === value) return done(current.revision);
                if (current.revision !== revisions.get(key)) {
                    return fail(storageError("StorageConflictError", "Another page has newer saved data; reload before saving"));
                }
                store.put({ value, revision: current.revision + 1 }, key);
                done(current.revision + 1);
            };
        }).then(revision => {
            revisions.set(key, revision);
            failures.delete(key);
        }));
        return job;
    }

    return {
        ready,
        getItem,
        setItem(key, value) {
            const text = String(value);
            if (!validState(key, text)) {
                const error = storageError("StorageDataError", "Invalid application state; existing data was preserved");
                report(error, key);
                const rejected = Promise.reject(error);
                rejected.catch(() => {});
                return rejected;
            }
            return write(key, text);
        },
        removeItem: key => write(key, null),
        clear() {
            return enqueue("clear", () => transactionResult(db, ["records"], "readwrite", (transaction, done, fail) => {
                const store = transaction.objectStore("records");
                const updates = new Map();
                for (const key of APP_STATE_KEYS) {
                    const request = store.get(key);
                    request.onsuccess = () => {
                        const current = request.result || { value: null, revision: 0 };
                        if (current.revision !== revisions.get(key)) {
                            return fail(storageError("StorageConflictError", "Another page has newer saved data; reload before clearing"));
                        }
                        updates.set(key, current.revision + 1);
                        if (updates.size !== APP_STATE_KEYS.length) return;
                        for (const [entry, revision] of updates) store.put({ value: null, revision }, entry);
                        done(updates);
                    };
                }
            }).then(updates => {
                for (const [key, revision] of updates) revisions.set(key, revision);
                failures.clear();
            }));
        },
        async flush() {
            await ready;
            await tail;
            if (failures.size) throw failures.values().next().value;
        },
        close() { closed = true; db?.close(); },
    };
}

export function createDeferredStateWriter(storage, delay = 50, serialize = value => value) {
    let pending = new Map();
    let timer = null;
    let active = Promise.resolve();
    function flush() {
        clearTimeout(timer);
        timer = null;
        const batch = pending;
        pending = new Map();
        active = active.catch(() => {}).then(async () => {
            for (const [key, value] of batch) await storage.setItem(key, serialize(value));
            await storage.flush();
        });
        active.catch(() => {});
        return active;
    }
    return {
        setItem(key, value) {
            pending.set(key, value);
            clearTimeout(timer);
            timer = setTimeout(flush, delay);
        },
        flush,
        removeItem(key) {
            pending.delete(key);
            active = active.catch(() => {}).then(() => storage.removeItem(key));
            active.catch(() => {});
            return active;
        },
    };
}

export function showStorageError(error) {
    if (typeof document === "undefined") return;
    const id = "laowu-storage-error";
    let notice = document.getElementById(id);
    if (!notice) {
        notice = document.createElement("div");
        notice.id = id;
        notice.setAttribute("role", "alert");
        notice.style.cssText = "position:fixed;z-index:2147483647;bottom:20px;left:50%;transform:translateX(-50%);width:max-content;max-width:calc(100vw - 32px);padding:12px 16px;border:1px solid #dc2626;border-radius:6px;background:#fff;color:#991b1b;font:14px/1.5 system-ui;box-shadow:0 4px 16px #0002";
        (document.body || document.documentElement).appendChild(notice);
    }
    notice.textContent = error?.name === "StorageConflictError"
        ? "\u5176\u4ed6\u9875\u9762\u5df2\u66f4\u65b0\u6570\u636e\uff0c\u672c\u9875\u6682\u505c\u8986\u76d6\u4fdd\u5b58\u3002\u8bf7\u5148\u4fdd\u5b58\u672c\u5730\u5de5\u7a0b\uff0c\u518d\u5237\u65b0\u9875\u9762\u3002"
        : "\u6d4f\u89c8\u5668\u5b58\u50a8\u8bfb\u5199\u5931\u8d25\uff0c\u65e7\u6570\u636e\u672a\u5220\u9664\u3002\u8bf7\u4fdd\u7559\u5f53\u524d\u9875\u9762\u5e76\u68c0\u67e5\u5b58\u50a8\u6743\u9650\u6216\u78c1\u76d8\u7a7a\u95f4\u3002";
}

export const appStateStorage = typeof window === "undefined"
    ? null
    : createAppStateStorage({ onError: showStorageError });
