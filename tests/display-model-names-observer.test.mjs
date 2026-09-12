import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

const source = await readFile(new URL("../display-model-names.js", import.meta.url), "utf8");

class ElementStub {
  constructor(selectors = [], text = "") {
    this.nodeType = 1;
    this.selectors = selectors;
    this.children = [];
    this.parentElement = null;
    this.isConnected = true;
    this.scans = 0;
    this.textNodes = text ? [{ nodeType: 3, nodeValue: text, parentElement: this }] : [];
  }

  append(child) {
    child.parentElement = this;
    this.children.push(child);
    return child;
  }

  matches(selector) {
    return selector.split(", ").some((part) => this.selectors.includes(part));
  }

  closest(selector) {
    return this.matches(selector) ? this : this.parentElement?.closest(selector) || null;
  }

  contains(node) {
    return node === this || this.children.some((child) => child.contains(node));
  }

  querySelectorAll(selector) {
    this.scans += 1;
    return this.children.flatMap((child) => [
      ...(child.matches(selector) ? [child] : []),
      ...child.querySelectorAll(selector),
    ]);
  }

  querySelector(selector) {
    return this.querySelectorAll(selector)[0] || null;
  }

  get textContent() {
    return this.textNodes.map((node) => node.nodeValue).join("")
      + this.children.map((child) => child.textContent).join("");
  }
}

function runtime() {
  let notify;
  let documentScans = 0;
  const frames = [];
  const document = {
    documentElement: {},
    querySelectorAll() {
      documentScans += 1;
      return [];
    },
    createTreeWalker(root) {
      const collect = (node) => [...node.textNodes, ...node.children.flatMap(collect)];
      const nodes = collect(root);
      let position = 0;
      return {
        currentNode: null,
        nextNode() {
          this.currentNode = nodes[position++] || null;
          return this.currentNode;
        },
      };
    },
  };
  const context = vm.createContext({
    document,
    Node: { ELEMENT_NODE: 1 },
    NodeFilter: { SHOW_TEXT: 4 },
    MutationObserver: class {
      constructor(callback) { notify = callback; }
      observe() {}
    },
    requestAnimationFrame(callback) {
      frames.push(callback);
      return frames.length;
    },
  });
  vm.runInContext(source, context);
  return {
    context,
    frames,
    notify: (mutations) => notify(mutations),
    documentScans: () => documentScans,
    flush() {
      const callbacks = frames.splice(0);
      callbacks.forEach((callback) => callback());
    },
  };
}

test("unrelated text mutations do not rescan the document or schedule work", () => {
  const app = runtime();
  assert.equal(app.documentScans(), 2);
  const prompt = new ElementStub([], "gpt-image-2-vip");
  app.notify(Array.from({ length: 500 }, () => ({
    type: "characterData",
    target: prompt.textNodes[0],
  })));
  assert.equal(app.frames.length, 0);
  assert.equal(app.documentScans(), 2);
  assert.equal(prompt.textContent, "gpt-image-2-vip");
});

test("repeated updates coalesce and normalize a model container itself", () => {
  const app = runtime();
  const picker = new ElementStub([".canvas-model-picker"], " gpt-image-2-vip ");
  app.notify(Array.from({ length: 100 }, () => ({
    type: "characterData",
    target: picker.textNodes[0],
  })));
  assert.equal(app.frames.length, 1);
  app.flush();
  assert.equal(picker.textContent, " gpt-image-2 ");
  assert.equal(picker.scans, 2);
  assert.equal(app.documentScans(), 2);
});

test("new model subtrees are processed without modifying neighboring prompt text", () => {
  const app = runtime();
  const wrapper = new ElementStub();
  const picker = wrapper.append(new ElementStub([".canvas-model-picker"], "gpt-image-2-vip"));
  const prompt = wrapper.append(new ElementStub([], "gpt-image-2-vip"));
  app.notify([{ type: "childList", target: new ElementStub(), addedNodes: [wrapper] }]);
  assert.equal(app.frames.length, 1);
  app.flush();
  assert.equal(picker.textContent, "gpt-image-2");
  assert.equal(prompt.textContent, "gpt-image-2-vip");
  assert.equal(app.documentScans(), 2);
});

test("queued ancestors subsume their descendants regardless of insertion order", () => {
  for (const parentFirst of [true, false]) {
    const app = runtime();
    const parent = new ElementStub([".canvas-model-menu"]);
    const child = parent.append(new ElementStub([".canvas-model-picker"], "gpt-image-2-vip"));
    for (const root of parentFirst ? [parent, child] : [child, parent]) {
      app.context.queueModelNormalization(root);
    }
    assert.equal(app.frames.length, 1);
    assert.equal(vm.runInContext("pendingModelRoots.size", app.context), 1);
    app.flush();
    assert.equal(child.textContent, "gpt-image-2");
  }
});

test("detached nodes are ignored before enqueue and before the frame runs", () => {
  const app = runtime();
  const picker = new ElementStub([".canvas-model-picker"], "gpt-image-2-vip");
  picker.isConnected = false;
  app.context.queueModelNormalization(picker);
  assert.equal(app.frames.length, 0);
  picker.isConnected = true;
  app.context.queueModelNormalization(picker);
  picker.isConnected = false;
  app.flush();
  assert.equal(picker.scans, 0);
  assert.equal(picker.textContent, "gpt-image-2-vip");
});
