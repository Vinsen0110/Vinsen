import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const start = bundle.indexOf("const rq=");
assert.notEqual(start, -1);
let componentSource;
for (let end = bundle.indexOf("}", start); end !== -1; end = bundle.indexOf("}", end + 1)) {
    const candidate = bundle.slice(start + "const rq=".length, end + 1);
    try {
        new vm.Script(`(${candidate})`);
        componentSource = candidate;
        break;
    } catch (error) {
        if (!(error instanceof SyntaxError)) throw error;
    }
}
assert.ok(componentSource, "extract the actual shipped theme component");

function harness(initialProps = {}, initialTheme = "light") {
    let cursor = 0, dirty = false, now = 0, nextTimer = 0, mounted = true;
    let props = initialProps;
    let dark = initialTheme === "dark";
    const hooks = [], observers = new Set(), effects = [], timers = new Map();
    const calls = { applied: [], errors: [], forbidden: [], afterUnmount: [] };
    const forbidden = name => () => {
        calls.forbidden.push(name);
        throw new Error(`${name} must not be called by theme toggling`);
    };
    const root = {
        classList: {
            contains: name => name === "dark" && dark,
            toggle(name, value) {
                assert.equal(name, "dark");
                dark = value;
                calls.applied.push(value ? "dark" : "light");
            },
        },
        style: { colorScheme: initialTheme },
        animate: forbidden("animate"),
        getBoundingClientRect: forbidden("getBoundingClientRect"),
    };
    const dependenciesChanged = (old, current) => !old || !current
        || old.length !== current.length || current.some((item, index) => !Object.is(item, old[index]));
    const context = vm.createContext({
        y: { jsx: (type, props) => ({ type, props }), jsxs: (type, props) => ({ type, props }) },
        no: (...values) => values.filter(Boolean).join(" "),
        $7: "DarkIcon",
        w7: "LightIcon",
        console: { error: (...args) => calls.errors.push(args) },
        document: { documentElement: root, startViewTransition: forbidden("startViewTransition") },
        window: { requestAnimationFrame: forbidden("requestAnimationFrame") },
        requestAnimationFrame: forbidden("requestAnimationFrame"),
        Lo: { flushSync: forbidden("flushSync") },
        setTimeout(callback, delay) {
            const id = ++nextTimer;
            timers.set(id, { callback, at: now + delay, delay });
            return id;
        },
        clearTimeout(id) { timers.delete(id); },
        MutationObserver: class {
            constructor(callback) { this.callback = callback; }
            observe() { observers.add(this); }
            disconnect() { observers.delete(this); }
        },
        c: {
            useState(initial) {
                const index = cursor++;
                if (!hooks[index]) hooks[index] = { value: typeof initial === "function" ? initial() : initial };
                return [hooks[index].value, value => {
                    if (!mounted) calls.afterUnmount.push(index);
                    const next = typeof value === "function" ? value(hooks[index].value) : value;
                    if (!Object.is(next, hooks[index].value)) dirty = true;
                    hooks[index].value = next;
                }];
            },
            useRef(initial) {
                const index = cursor++;
                if (!hooks[index]) hooks[index] = { current: initial };
                return hooks[index];
            },
            useEffect(effect, deps) {
                const index = cursor++;
                const previous = hooks[index];
                if (dependenciesChanged(previous?.deps, deps)) {
                    hooks[index] = { deps, cleanup: previous?.cleanup };
                    effects.push(() => {
                        hooks[index].cleanup?.();
                        hooks[index].cleanup = effect();
                    });
                }
            },
            useCallback(callback, deps) {
                const index = cursor++;
                if (dependenciesChanged(hooks[index]?.deps, deps)) hooks[index] = { deps, callback };
                return hooks[index].callback;
            },
        },
    });
    const component = vm.runInContext(`(${componentSource})`, context);
    const render = nextProps => {
        if (nextProps) props = nextProps;
        let tree, passes = 0;
        do {
            dirty = false;
            cursor = 0;
            tree = component(props);
            while (effects.length) effects.shift()();
            assert.ok(++passes < 10, "theme effects must settle without a render loop");
        } while (dirty);
        return tree;
    };
    const advance = milliseconds => {
        const target = now + milliseconds;
        for (;;) {
            const due = [...timers.entries()].filter(([, timer]) => timer.at <= target)
                .sort((a, b) => a[1].at - b[1].at)[0];
            if (!due) break;
            now = due[1].at;
            timers.delete(due[0]);
            due[1].callback();
        }
        now = target;
    };
    return {
        render, advance, timers, calls, root, observers,
        theme: () => dark ? "dark" : "light",
        externalTheme(theme) {
            dark = theme === "dark";
            root.style.colorScheme = theme;
            observers.forEach(observer => observer.callback());
        },
        unmount() {
            mounted = false;
            hooks.forEach(hook => hook?.cleanup?.());
        },
    };
}

test("rapid clicks on the same rendered handler apply one theme and expose disabled/busy state", () => {
    const changes = [];
    const h = harness({ onThemeChange: value => changes.push(value) });
    const button = h.render();
    for (let index = 0; index < 20; index += 1) button.props.onClick();
    assert.deepEqual(changes, ["dark"]);
    assert.deepEqual(h.calls.applied, ["dark"]);
    const busy = h.render();
    assert.equal(busy.props.disabled, true);
    assert.equal(busy.props["aria-busy"], true);
    assert.equal(h.root.style.colorScheme, "dark");
    assert.equal(h.timers.size, 1);
    assert.equal([...h.timers.values()][0].delay, 180);
    assert.deepEqual(h.calls.forbidden, []);
});

test("after cooldown the same stale handler reads the actual DOM theme and switches back", () => {
    const changes = [];
    const h = harness({ theme: "light", onThemeChange: value => changes.push(value) });
    const button = h.render();
    button.props.onClick();
    h.advance(179);
    button.props.onClick();
    assert.deepEqual(changes, ["dark"]);
    h.advance(1);
    assert.equal(h.render().props.disabled, false);
    assert.equal(h.render().props["aria-busy"], false);
    button.props.onClick();
    assert.deepEqual(changes, ["dark", "light"]);
    assert.equal(h.theme(), "light");
    assert.deepEqual(h.calls.forbidden, []);
});

test("unmount clears the cooldown timer and DOM observer without later state updates", () => {
    const h = harness();
    h.render().props.onClick();
    assert.equal(h.timers.size, 1);
    assert.equal(h.observers.size, 1);
    h.unmount();
    assert.equal(h.timers.size, 0);
    assert.equal(h.observers.size, 0);
    h.advance(1_000);
    assert.deepEqual(h.calls.afterUnmount, []);
});

test("spread props cannot override the internal click handler or computed disabled/busy attributes", () => {
    const changes = [];
    let bypass = 0;
    const h = harness({
        onThemeChange: value => changes.push(value),
        onClick: () => { bypass += 1; },
        disabled: false,
        "aria-busy": false,
        "aria-label": "Theme",
    });
    const button = h.render();
    button.props.onClick();
    assert.deepEqual(changes, ["dark"]);
    assert.equal(bypass, 0);
    const busy = h.render();
    assert.equal(busy.props.disabled, true);
    assert.equal(busy.props["aria-busy"], true);
    assert.equal(busy.props["aria-label"], "Theme");
});

test("an externally disabled button cannot mutate the theme even if its handler is called", () => {
    const changes = [];
    const h = harness({ disabled: true, onThemeChange: value => changes.push(value) });
    h.render().props.onClick();
    assert.equal(h.theme(), "light");
    assert.equal(h.timers.size, 0);
    assert.deepEqual(changes, []);
});

test("external theme changes synchronize the uncontrolled icon and subsequent toggle direction", () => {
    const changes = [];
    const h = harness({ onThemeChange: value => changes.push(value) });
    assert.equal(h.render().props.children[0].type, "LightIcon");
    h.externalTheme("dark");
    assert.equal(h.render().props.children[0].type, "DarkIcon");
    h.render().props.onClick();
    assert.deepEqual(changes, ["light"]);
    assert.equal(h.theme(), "light");
});

test("controlled theme changes update the icon without creating a DOM observer", () => {
    const h = harness({ theme: "light" });
    assert.equal(h.render().props.children[0].type, "LightIcon");
    h.externalTheme("dark");
    assert.equal(h.render({ theme: "dark" }).props.children[0].type, "DarkIcon");
    assert.equal(h.observers.size, 0);
});

test("an already active target theme does not write the store or enter cooldown", () => {
    const changes = [];
    const h = harness({ targetTheme: "dark", onThemeChange: value => changes.push(value) }, "dark");
    h.render().props.onClick();
    assert.deepEqual(changes, []);
    assert.deepEqual(h.calls.applied, []);
    assert.equal(h.timers.size, 0);
    assert.equal(h.render().props.disabled, false);
});

test("a store callback error is contained and cannot leave the theme button permanently locked", () => {
    const h = harness({ onThemeChange() { throw new Error("store failure"); } });
    const button = h.render();
    assert.doesNotThrow(() => button.props.onClick());
    assert.equal(h.calls.errors.length, 1);
    assert.equal(h.render().props.disabled, true);
    h.advance(180);
    assert.equal(h.render().props.disabled, false);
    assert.doesNotThrow(() => button.props.onClick());
    assert.equal(h.calls.errors.length, 2);
    assert.deepEqual(h.calls.applied, ["dark", "light"]);
});

test("the actual canvas toolbar passes the persistent store setter to the theme button", () => {
    const changes = [];
    const setTheme = theme => changes.push(theme);
    const wrapperStart = bundle.indexOf("function o$e(");
    const wrapperEnd = bundle.indexOf("function ", wrapperStart + 9);
    const jsx = (type, props) => ({ type, props });
    const context = vm.createContext({
        fr: selector => selector({ theme: "light", setTheme }),
        Xr: selector => selector({ openConfigDialog() {} }),
        xr: { light: { node: { text: "#111" } } },
        rq: "ThemeToggle",
        y: { jsx, jsxs: jsx },
    });
    vm.runInContext(bundle.slice(wrapperStart, wrapperEnd), context);
    const wrapper = context.o$e({ showConfig: false, variant: "canvas" });
    const themeProps = wrapper.props.children.find(child => child?.type === "ThemeToggle").props;
    assert.equal(themeProps.onThemeChange, setTheme);
    const h = harness(themeProps);
    h.render().props.onClick();
    assert.deepEqual(changes, ["dark"]);
    assert.equal(h.root.style.colorScheme, "dark");
});

test("the theme control has no snapshot animation, synchronous React flush, or extra animation frame", () => {
    assert.doesNotMatch(componentSource, /startViewTransition|\.animate\(|clipPath|flushSync|requestAnimationFrame/);
});
