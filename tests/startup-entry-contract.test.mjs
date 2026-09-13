import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const html = await readFile(new URL("../index.html", import.meta.url), "utf8");
const start = bundle.indexOf("function Dke(");
const end = bundle.indexOf("function J6(", start);
assert.ok(start >= 0 && end > start, "startup component boundaries must remain identifiable");

const jsx = (type, props) => ({ type, props });
const renderStartup = vm.runInNewContext(`(${bundle.slice(start, end)})`, {
    y: { jsx, jsxs: jsx },
    Bt: "Button",
    BP: "NewProjectIcon",
    x5: "OpenProjectIcon",
});

function descendants(node) {
    if (Array.isArray(node)) return node.flatMap(descendants);
    if (!node || typeof node !== "object") return [];
    return [node, ...descendants(node.props?.children)];
}

function textContent(node) {
    if (Array.isArray(node)) return node.map(textContent).join("");
    if (node == null || typeof node === "boolean") return "";
    return typeof node === "object" ? textContent(node.props?.children) : String(node);
}

function startup(props = {}) {
    return renderStartup({ busy: false, error: "", onOpenFolder() {}, ...props });
}

function previewPicker() {
    const script = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)]
        .map((match) => match[1])
        .find((source) => source.includes('var pickerMode = "blank"'));
    assert.ok(script, "browser preview picker shim must exist");
    let clickHandler;
    let inputClicks = 0;
    const window = {};
    const document = {
        addEventListener(type, listener) {
            if (type === "click") clickHandler = listener;
        },
        body: { appendChild() {} },
        createElement(type) {
            assert.equal(type, "input");
            const handlers = {};
            return {
                files: [],
                addEventListener(event, listener) { handlers[event] = listener; },
                click() {
                    inputClicks += 1;
                    handlers.cancel();
                },
                remove() {},
            };
        },
    };
    vm.runInNewContext(script, {
        window, document, URLSearchParams, DOMException,
        location: { search: "?preview=1" },
    });
    assert.equal(typeof clickHandler, "function");
    return {
        window,
        inputClicks: () => inputClicks,
        click(label, action) {
            const button = {
                textContent: label,
                dataset: { projectAction: action },
                getAttribute(name) { return name === "data-project-action" ? action ?? null : null; },
            };
            clickHandler({ target: { closest: () => button } });
        },
    };
}

test("startup entry renders the brand and its local image", () => {
    const root = startup();
    const nodes = descendants(root);
    assert.equal(root.type, "main");
    assert.match(root.props.className, /(?:^|\s)startup-entry(?:\s|$)/);
    assert.equal(textContent(nodes.find((node) => node.type === "h1")), "\u8001\u5c4b\u4f5c\u574a");
    const logo = nodes.find((node) => node.type === "img" && node.props.src === "./kitty.png");
    assert.ok(logo, "startup branding must use the existing local image");
    assert.equal(typeof logo.props.alt, "string", "brand image needs an explicit alt attribute");
});

test("startup actions preserve blank and import callbacks with explicit picker modes", () => {
    const calls = [];
    const buttons = descendants(startup({ onOpenFolder: (mode) => calls.push(mode) }))
        .filter((node) => node.type === "Button");
    assert.equal(buttons.length, 2);
    const blank = buttons.find((node) => node.props["data-project-action"] === "blank");
    const open = buttons.find((node) => node.props["data-project-action"] === "import");
    assert.ok(blank);
    assert.ok(open);
    assert.equal(textContent(blank), "\u65b0\u5efa\u7a7a\u767d\u5de5\u7a0b");
    assert.equal(textContent(open), "\u6253\u5f00\u5df2\u6709\u5de5\u7a0b");
    blank.props.onClick();
    open.props.onClick();
    assert.deepEqual(calls, ["blank", "import"]);
});

test("startup keeps busy feedback and exposes project errors unchanged", () => {
    const error = "Project directory permission was denied";
    const root = startup({ busy: true, error });
    const nodes = descendants(root);
    const blank = nodes.find((node) => node.props["data-project-action"] === "blank");
    const open = nodes.find((node) => node.props["data-project-action"] === "import");
    assert.ok(blank);
    assert.ok(open);
    assert.equal(blank.props.loading, true);
    assert.equal(open.props.disabled, true);
    assert.equal(nodes.filter((node) => node.props.children === error).length, 1);
    const idle = descendants(startup());
    assert.equal(idle.find((node) => node.props["data-project-action"] === "blank").props.loading, false);
    assert.equal(idle.find((node) => node.props["data-project-action"] === "import").props.disabled, false);
    assert.ok(!textContent(startup()).includes(error));
});

test("startup stylesheet is loaded and supplies a scoped dark theme", async () => {
    assert.ok(
        /<link\b(?=[^>]*\brel=["']stylesheet["'])(?=[^>]*\bhref=["'][^"']*startup-entry\.css(?:\?[^"']*)?["'])[^>]*>/.test(html),
        "index.html must load the startup entry stylesheet",
    );
    const css = await readFile(new URL("../startup-entry.css", import.meta.url), "utf8");
    assert.match(css, /\.startup-entry\b/);
    assert.match(css, /\.dark\s+\.startup-entry\b/);
});

test("theme hydration restores saved dark mode and defaults invalid values to light", () => {
    const themeStart = bundle.indexOf("const fr=");
    const themeEnd = bundle.indexOf("Pxe=", themeStart);
    assert.ok(themeStart >= 0 && themeEnd > themeStart, "theme store boundaries must remain identifiable");
    const expression = bundle.slice(themeStart + "const fr=".length, themeEnd).trim().replace(/,$/, "");
    const store = vm.runInNewContext(`(${expression})`, {
        fh: () => (config) => config,
        SS: (initialize, options) => ({ initialize, options }),
    });
    const state = store.initialize(() => {});
    assert.equal(state.theme, "light");
    assert.equal(store.options.name, "infinite-canvas:theme_store");
    assert.equal(store.options.skipHydration, true);
    for (const [persisted, expected] of [
        [{ theme: "dark" }, "dark"],
        [{ theme: "light" }, "light"],
        [undefined, "light"],
        [null, "light"],
        [{}, "light"],
        [{ theme: "system" }, "light"],
        [{ theme: "" }, "light"],
        [{ theme: 1 }, "light"],
    ]) {
        const merged = store.options.merge(persisted, state);
        assert.equal(merged.theme, expected, `unexpected restored theme for ${JSON.stringify(persisted)}`);
        assert.equal(merged.setTheme, state.setTheme, "hydration must preserve the store action");
    }
});

test("preview picker uses explicit action instead of the renamed button label", async () => {
    const picker = previewPicker();
    picker.click("\u6253\u5f00\u5df2\u6709\u5de5\u7a0b", "import");
    await assert.rejects(picker.window.showDirectoryPicker(), { name: "AbortError" });
    assert.equal(picker.inputClicks(), 1);
    picker.click("\u5bfc\u5165", "blank");
    const directory = await picker.window.showDirectoryPicker();
    assert.equal(typeof directory.getFileHandle, "function");
    assert.equal(picker.inputClicks(), 1, "new project must not open the import file picker");
});

test("preview picker retains legacy label fallback for existing buttons", async () => {
    const picker = previewPicker();
    picker.click("\u5bfc\u5165\u672c\u5730\u5de5\u7a0b");
    await assert.rejects(picker.window.showDirectoryPicker(), { name: "AbortError" });
    picker.click("\u65b0\u5efa\u7a7a\u767d\u5de5\u7a0b");
    const directory = await picker.window.showDirectoryPicker();
    assert.equal(typeof directory.getDirectoryHandle, "function");
    assert.equal(picker.inputClicks(), 1);
});
