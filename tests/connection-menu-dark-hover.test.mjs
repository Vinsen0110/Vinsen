import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const html = await readFile(new URL("../index.html", import.meta.url), "utf8");
const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const css = [...html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/g)]
    .map(match => match[1]).join("\n").replace(/\/\*[\s\S]*?\*\//g, "");
const menuRules = [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)]
    .filter(match => match[1].includes("[data-connection-create-menu]"))
    .map(match => ({ selectors: match[1].split(",").map(selector => selector.trim()), body: match[2] }));

test("connection menu hover and keyboard focus use a scoped dark surface and readable text", () => {
    for (const state of ["hover", "focus-visible"]) {
        const rule = menuRules.find(entry => entry.selectors.some(selector =>
            selector.endsWith(`[data-connection-create-menu] button:${state}`)));
        assert.ok(rule, `the menu must style button:${state}`);
        assert.match(rule.body, /background(?:-color)?\s*:\s*#3a3631\s*!important\s*;/);
        assert.match(rule.body, /color\s*:\s*#f5f5f4\s*!important\s*;/);
    }
});

test("connection menu hover overrides cannot apply to light mode or recolor child icons", () => {
    assert.ok(menuRules.length > 0);
    for (const rule of menuRules) {
        for (const selector of rule.selectors) {
            assert.match(selector, /^(?:html)?\.dark\s+\[data-connection-create-menu\]\s+button(?::(?:hover|focus-visible))?$/);
            assert.doesNotMatch(selector, /span|svg|\*/);
        }
    }
});

function extracted(name) {
    const start = bundle.indexOf(`function ${name}(`);
    assert.notEqual(start, -1);
    const end = bundle.indexOf("function ", start + 9);
    return bundle.slice(start, end);
}

function runtime() {
    const jsx = (type, props) => ({ type, props });
    const context = vm.createContext({
        y: { jsx, jsxs: jsx },
        c: {
            useRef: current => ({ current }),
            useState: initial => [initial, () => {}],
            useLayoutEffect: () => {},
        },
        Lo: { createPortal: element => element },
        document: { body: {} },
        window: { innerHeight: 900 },
        xr: { dark: { node: { stroke: "#44403c" } } },
        fr: selector => selector({ theme: "dark" }),
        g7: "ImageIcon",
        Is: "GenerateIcon",
    });
    vm.runInContext(extracted("eB"), context);
    vm.runInContext(extracted("zke"), context);
    return context;
}

function descendants(tree) {
    if (!tree || typeof tree !== "object") return [];
    if (Array.isArray(tree)) return tree.flatMap(descendants);
    const children = typeof tree.type === "function" ? tree.type(tree.props) : tree.props.children;
    return [tree, ...descendants(children)];
}

test("the real eB button keeps its callback identity and semantic icon color", () => {
    const context = runtime();
    let clicks = 0;
    const onClick = () => { clicks += 1; };
    const icon = { type: "CustomIcon" };
    const button = context.eB({ icon, title: "Image", color: "#ff4fa3", onClick });
    assert.equal(button.props.onClick, onClick);
    assert.equal(button.props.children[0].props.style.color, "#ff4fa3");
    assert.equal(button.props.children[0].props.children, icon);
    assert.match(button.props.className, /hover:bg-stone-100/);
    button.props.onClick();
    assert.equal(clicks, 1);
});

test("the real connection menu keeps all create, close and canvas propagation callbacks", () => {
    const context = runtime();
    const created = [];
    let closed = 0, stopped = 0;
    const menu = context.zke({
        pending: { position: { x: 100, y: 200 } },
        onCreate: kind => created.push(kind),
        onClose: () => { closed += 1; },
    });
    assert.equal(menu.props["data-connection-create-menu"], true);
    const buttons = descendants(menu).filter(node => node.type === "button");
    assert.equal(buttons.length, 4);
    buttons.forEach(button => button.props.onClick());
    assert.deepEqual(created, ["image-input", "text-generation", "image-generation"]);
    assert.equal(closed, 1);
    menu.props.onPointerDown({ stopPropagation: () => { stopped += 1; } });
    menu.props.onMouseDown({ stopPropagation: () => { stopped += 1; } });
    assert.equal(stopped, 2);
    assert.deepEqual(
        buttons.slice(1).map(button => button.props.children[0].props.style.color),
        ["#2f80ff", "#22c55e", "#ff4fa3"],
    );
});
