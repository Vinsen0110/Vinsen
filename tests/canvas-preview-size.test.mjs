import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

function extractFunction(name) {
    const start = bundle.indexOf(`function ${name}(`);
    assert.notEqual(start, -1, `${name} must exist in the browser bundle`);
    // Let the parser find the complete function, including nested object braces.
    for (let end = bundle.indexOf("}", start); end !== -1; end = bundle.indexOf("}", end + 1)) {
        let script;
        try {
            script = new vm.Script(`(${bundle.slice(start, end + 1)})`);
        } catch (error) {
            if (error instanceof SyntaxError) continue;
            throw error;
        }
        return script.runInNewContext({}, { timeout: 1000 });
    }
    assert.fail(`could not extract ${name}`);
}

const cases = [
    ["tall portraits use a 1280px longest edge", 1000, 4000, 320, 1280],
    ["landscapes use a 1280px longest edge", 4000, 1000, 1280, 320],
    ["4K portraits retain their aspect ratio", 2160, 3840, 720, 1280],
    ["large squares become 1280px squares", 4096, 4096, 1280, 1280],
    ["small images are not enlarged", 320, 640, 320, 640],
    ["extremely narrow images keep at least one pixel", 1, 4000, 1, 1280],
];

for (const [name, width, height, expectedWidth, expectedHeight] of cases) {
    test(name, () => {
        const dimensions = extractFunction("canvasPreviewDimensions");
        assert.deepEqual(
            { ...dimensions(width, height) },
            { width: expectedWidth, height: expectedHeight },
        );
    });
}

for (const value of [0, NaN, Infinity, -1]) {
    test(`preview dimensions reject ${String(value)} in either dimension`, () => {
        const dimensions = extractFunction("canvasPreviewDimensions");
        assert.throws(() => dimensions(value, 4000));
        assert.throws(() => dimensions(4000, value));
    });
}

test("persistent preview keys use v2 so older oversized previews are not reused", () => {
    const persistentKey = extractFunction("canvasPreviewPersistentKey");
    assert.equal(persistentKey("image:portrait-test"), "preview-v2:image:portrait-test");
    assert.equal(persistentKey("image:landscape-test"), "preview-v2:image:landscape-test");
});
