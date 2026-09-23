export const GROUP_COLORS = ["#22a06b", "#4285e8", "#d59b22", "#d75a87", "#8b6bc9", "#718096"];
const DEFAULT_COLOR = GROUP_COLORS[0];
const cleanName = value => String(value || "").trim().slice(0, 80) || "\u5206\u7ec4";
const cleanColor = value => /^#[\da-f]{6}$/i.test(value || "") ? value : DEFAULT_COLOR;
const validFrame = frame => frame && ["x", "y", "width", "height"].every(key => Number.isFinite(frame[key])) &&
    frame.width >= 240 && frame.height >= 100;

export function canvasGroupOf(node) {
    const group = node?.canvasGroup;
    return group && typeof group.id === "string" && group.id
        ? { id: group.id, name: cleanName(group.name), color: cleanColor(group.color),
            ...(validFrame(group.frame) ? { frame: { ...group.frame } } : {}) }
        : null;
}

export function collectCanvasGroups(nodes) {
    const groups = new Map();
    for (const node of nodes) {
        const value = canvasGroupOf(node);
        if (!value) continue;
        if (!groups.has(value.id)) groups.set(value.id, { ...value, nodes: [] });
        groups.get(value.id).nodes.push(node);
    }
    return [...groups.values()];
}

export function canvasGroupBounds(nodes) {
    if (!nodes.length) return null;
    const frame = canvasGroupOf(nodes[0])?.frame;
    if (frame) return frame;
    let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
    for (const node of nodes) {
        left = Math.min(left, node.position.x);
        top = Math.min(top, node.position.y);
        right = Math.max(right, node.position.x + node.width);
        bottom = Math.max(bottom, node.position.y + node.height);
    }
    return { x: left - 24, y: top - 64, width: Math.max(240, right - left + 48), height: bottom - top + 88 };
}

// Keep membership on the node, outside model metadata, so existing history and
// project serializers preserve it without sending it to generation providers.
export function assignCanvasGroup(nodes, ids, group) {
    const selected = new Set(ids);
    let changed = false;
    const next = nodes.map(node => {
        if (!selected.has(node.id)) return node;
        const previous = canvasGroupOf(node);
        const value = group ? canvasGroupOf({ canvasGroup: group }) : null;
        if (!value && !node.canvasGroup ||
            value && previous?.id === value.id && previous.name === value.name && previous.color === value.color &&
            JSON.stringify(previous.frame) === JSON.stringify(value.frame)) return node;
        changed = true;
        const result = { ...node };
        if (value) result.canvasGroup = { ...value, ...(value.frame ? { frameOrigin: { ...node.position } } : {}) };
        else delete result.canvasGroup;
        return result;
    });
    return changed ? next : nodes;
}

export function createCanvasGroup(nodes, ids, createId = () => crypto.randomUUID()) {
    const selected = new Set(ids);
    if (nodes.filter(node => selected.has(node.id)).length < 2) return nodes;
    const names = new Set(collectCanvasGroups(nodes).map(group => group.name));
    let number = 1;
    while (names.has(`\u5206\u7ec4 ${number}`)) number++;
    return assignCanvasGroup(nodes, selected, {
        id: createId(), name: `\u5206\u7ec4 ${number}`, color: DEFAULT_COLOR,
    });
}

export function updateCanvasGroup(nodes, id, changes) {
    const group = collectCanvasGroups(nodes).find(item => item.id === id);
    if (!group) return nodes;
    return assignCanvasGroup(nodes, group.nodes.map(node => node.id), { ...group, ...changes, id });
}

export function dissolveCanvasGroup(nodes, id) {
    return assignCanvasGroup(nodes, nodes.filter(node => canvasGroupOf(node)?.id === id).map(node => node.id), null);
}

export function remapClonedCanvasGroups(nodes, createId = () => crypto.randomUUID()) {
    let result = nodes;
    for (const group of collectCanvasGroups(nodes)) {
        const origin = group.nodes[0].canvasGroup.frameOrigin;
        const position = group.nodes[0].position;
        const frame = group.frame && origin ? { ...group.frame,
            x: group.frame.x + position.x - origin.x, y: group.frame.y + position.y - origin.y } : group.frame;
        result = assignCanvasGroup(result, group.nodes.map(node => node.id),
            group.nodes.length > 1 ? { ...group, frame, id: createId() } : null);
    }
    return result;
}

const insideGroupFrame = (node, frame) => {
    const x = node.position.x + node.width / 2, y = node.position.y + node.height / 2;
    return x >= frame.x && x <= frame.x + frame.width && y >= frame.y && y <= frame.y + frame.height;
};

export function placeCanvasGroupClones(copies, sources, scene, createId) {
    const groups = new Map(collectCanvasGroups(scene).map(group => [group.id, group]));
    const liveNodes = new Map(scene.map(node => [node.id, node]));
    const targets = new Map();
    const prepared = copies.map((copy, index) => {
        const source = sources[index], savedGroup = canvasGroupOf(source);
        if (!savedGroup) {
            if (!copy.canvasGroup) return copy;
            const { canvasGroup, ...node } = copy;
            return node;
        }
        const group = groups.get(savedGroup.id), liveSource = liveNodes.get(source.id);
        if (group) {
            const frame = canvasGroupBounds(group.nodes);
            // Membership alone is insufficient after a member has moved outside
            // a manually sized frame. Never absorb unrelated outside nodes.
            if (!liveSource || canvasGroupOf(liveSource)?.id !== group.id || !insideGroupFrame(liveSource, frame)) {
                const { canvasGroup, ...node } = copy;
                return node;
            }
            if (insideGroupFrame(copy, frame)) targets.set(copy.id, group);
        }
        return { ...copy, canvasGroup: { ...source.canvasGroup } };
    });
    const outside = prepared.filter(node => !targets.has(node.id));
    const independent = new Map(remapClonedCanvasGroups(outside, createId).map(node => [node.id, node]));
    return prepared.map(node => {
        const group = targets.get(node.id);
        return group ? { ...node, canvasGroup: {
            ...canvasGroupOf({ canvasGroup: group }),
            ...(group.frame ? { frameOrigin: { ...node.position } } : {}),
        } } : independent.get(node.id);
    });
}

export function finishCanvasGroupCloneDrag(nodes, context) {
    if (!context) return nodes;
    const byId = new Map(nodes.map(node => [node.id, node]));
    const copies = [], sources = [];
    context.copyIds.forEach((id, index) => {
        if (byId.has(id)) { copies.push(byId.get(id)); sources.push(context.sources[index]); }
    });
    // Use the pre-copy scene so automatic bounds cannot expand to include a
    // dragged copy before its final drop position has been checked.
    const placed = new Map(placeCanvasGroupClones(copies, sources, context.scene).map(node => [node.id, node]));
    return nodes.map(node => placed.get(node.id) || node);
}

// Translate a manual frame only when all of its members move together.
// Individual member movement leaves the frame and the other members untouched.
export function moveCanvasGroupFrames(before, after) {
    if (!after.some(node => validFrame(node.canvasGroup?.frame))) return after;
    const previous = new Map(before.map(node => [node.id, node]));
    const frames = new Map();
    for (const group of collectCanvasGroups(after)) {
        if (!group.frame) continue;
        const deltas = group.nodes.map(node => {
            const old = previous.get(node.id);
            return old ? { x: node.position.x - old.position.x, y: node.position.y - old.position.y } : null;
        });
        const delta = deltas[0];
        const together = delta && deltas.every(item => item && Math.abs(item.x - delta.x) < 1e-6 && Math.abs(item.y - delta.y) < 1e-6);
        frames.set(group.id, together ? { ...group.frame, x: group.frame.x + delta.x, y: group.frame.y + delta.y } : group.frame);
    }
    if (!frames.size) return after;
    return after.map(node => {
        const frame = frames.get(node.canvasGroup?.id);
        if (!frame) return node;
        const previousFrame = node.canvasGroup.frame, origin = node.canvasGroup.frameOrigin;
        if (["x", "y", "width", "height"].every(key => previousFrame[key] === frame[key]) &&
            origin?.x === node.position.x && origin?.y === node.position.y) return node;
        return { ...node, canvasGroup: { ...node.canvasGroup, frame, frameOrigin: { ...node.position } } };
    });
}

export function resizeCanvasGroupFrame(frame, direction, dx, dy) {
    let { x, y, width, height } = frame;
    if (direction.includes("e")) width = Math.max(240, width + dx);
    if (direction.includes("s")) height = Math.max(100, height + dy);
    if (direction.includes("w")) { width = Math.max(240, width - dx); x += frame.width - width; }
    if (direction.includes("n")) { height = Math.max(100, height - dy); y += frame.height - height; }
    return { x, y, width, height };
}

export function beginCanvasGroupResize(event, group, direction, scale, commit, onEnd, host = window) {
    const element = event.currentTarget.closest(".canvas-group-frame");
    if (!element) return () => {};
    event.preventDefault();
    event.stopPropagation();
    const original = canvasGroupBounds(group.nodes);
    let pending = original, raf = null, ended = false;
    const paint = frame => {
        for (const key of ["x", "y", "width", "height"]) {
            element.style[key === "x" ? "left" : key === "y" ? "top" : key] = `${frame[key]}px`;
        }
    };
    const update = next => {
        pending = resizeCanvasGroupFrame(original, direction,
            (next.clientX - event.clientX) / scale, (next.clientY - event.clientY) / scale);
    };
    const move = next => {
        if (next.pointerId !== event.pointerId || !(next.buttons & 1)) return;
        next.preventDefault();
        update(next);
        if (raf === null) raf = host.requestAnimationFrame(() => { raf = null; paint(pending); });
    };
    const finish = (next, cancel = false) => {
        if (ended || next?.pointerId != null && next.pointerId !== event.pointerId) return;
        if (next && !cancel && next.button !== 0) return;
        ended = true;
        if (raf !== null) host.cancelAnimationFrame(raf);
        host.removeEventListener("pointermove", move, true);
        host.removeEventListener("pointerup", finish, true);
        host.removeEventListener("pointercancel", cancelPointer, true);
        host.removeEventListener("blur", cancelAll);
        host.removeEventListener("keydown", keydown, true);
        element.removeAttribute("data-group-resizing");
        if (cancel) paint(original);
        else {
            update(next);
            paint(pending);
            if (JSON.stringify(pending) !== JSON.stringify(original)) commit(pending);
        }
        onEnd();
    };
    const cancelPointer = next => finish(next, true);
    const cancelAll = () => finish(null, true);
    const keydown = next => {
        if (next.key === "Escape") { next.preventDefault(); next.stopImmediatePropagation(); cancelAll(); }
        else if (next.key === "Delete" || next.key === "Backspace" || next.ctrlKey || next.metaKey || next.altKey) {
            next.preventDefault(); next.stopImmediatePropagation();
        }
    };
    element.setAttribute("data-group-resizing", "true");
    host.addEventListener("pointermove", move, { capture: true, passive: false });
    host.addEventListener("pointerup", finish, true);
    host.addEventListener("pointercancel", cancelPointer, true);
    host.addEventListener("blur", cancelAll);
    host.addEventListener("keydown", keydown, true);
    return cancelAll;
}

export function isCanvasGroupShortcut(event) {
    return event.code === "KeyG" && event.altKey && !event.ctrlKey && !event.metaKey &&
        !event.shiftKey && !event.repeat && !event.isComposing && !event.defaultPrevented;
}

export function useCanvasGroups(React, options) {
    const { nodes, selected, selectedRef, setSelected, startDrag, canInteract } = options;
    const latest = React.useRef(options);
    latest.current = options;
    const groupSelection = React.useRef(null);
    const resizing = React.useRef(null);
    React.useEffect(() => () => resizing.current?.(), []);
    const [menu, setMenu] = React.useState(null);
    const [editing, setEditing] = React.useState(null);
    const groups = React.useMemo(() => collectCanvasGroups(nodes), [nodes]);
    const change = transform => latest.current.commit(transform(latest.current.nodesRef.current));
    const actionIds = () => menu?.nodeIds || latest.current.selectedRef.current;
    const createFromSelection = ids => {
        const current = latest.current;
        const next = createCanvasGroup(current.nodesRef.current, ids);
        if (next === current.nodesRef.current) return;
        current.commit(next);
        const selectedIds = new Set(ids);
        groupSelection.current = selectedIds;
        current.selectedRef.current = selectedIds;
        current.setSelected(selectedIds);
    };
    const create = () => createFromSelection(actionIds());
    const dissolve = id => {
        change(current => dissolveCanvasGroup(current, id));
        groupSelection.current = null;
        const ids = new Set();
        latest.current.selectedRef.current = ids;
        latest.current.setSelected(ids);
        setMenu(null);
    };
    React.useEffect(() => {
        const keydown = event => {
            if (event.key === "Escape") { setMenu(null); setEditing(null); }
            if (resizing.current || !latest.current.canInteract()) return;
            if (event.target?.closest?.("input,textarea,select,[contenteditable]:not([contenteditable='false']),[data-canvas-no-zoom],[role='dialog'],[role='menu'],.ant-modal,.ant-select-dropdown")) return;
            if ((event.key === "Delete" || event.key === "Backspace") && !event.ctrlKey && !event.metaKey &&
                !event.altKey && !event.shiftKey && groupSelection.current === latest.current.selectedRef.current) {
                const ids = latest.current.selectedRef.current;
                const group = collectCanvasGroups(latest.current.nodesRef.current)
                    .find(item => item.nodes.length === ids.size && item.nodes.every(node => ids.has(node.id)));
                if (group) {
                    event.preventDefault();
                    event.stopImmediatePropagation();
                    dissolve(group.id);
                    return;
                }
            }
            if (!isCanvasGroupShortcut(event)) return;
            event.preventDefault();
            createFromSelection(latest.current.selectedRef.current);
        };
        window.addEventListener("keydown", keydown, true);
        return () => window.removeEventListener("keydown", keydown, true);
    }, []);
    React.useEffect(() => {
        if (!menu) return;
        const close = event => {
            if (!event.target?.closest?.("[data-canvas-group-menu]")) setMenu(null);
        };
        window.addEventListener("pointerdown", close, true);
        window.addEventListener("blur", close);
        window.addEventListener("resize", close);
        window.addEventListener("wheel", close, { passive: true });
        return () => {
            window.removeEventListener("pointerdown", close, true);
            window.removeEventListener("blur", close);
            window.removeEventListener("resize", close);
            window.removeEventListener("wheel", close);
        };
    }, [menu]);
    React.useEffect(() => { setMenu(null); }, [options.viewport]);
    const selectGroup = group => {
        const ids = new Set(group.nodes.map(node => node.id));
        groupSelection.current = ids;
        selectedRef.current = ids;
        setSelected(ids);
    };
    const openMenu = (event, groupId, nodeId, palette = false) => {
        event.preventDefault();
        event.stopPropagation();
        if (nodeId && !selectedRef.current.has(nodeId)) {
            const ids = new Set([nodeId]);
            selectedRef.current = ids;
            setSelected(ids);
        }
        setMenu({ x: event.clientX, y: event.clientY, groupId, palette, nodeIds: [...selectedRef.current] });
    };
    return {
        groups, selected, menu, editing, setEditing, setMenu, selectGroup, openMenu, create,
        openNodeMenu: (event, id) => openMenu(event, null, id),
        beforeNodeDrag: (event, id) => {
            if (event.button !== 0) return;
            const current = latest.current;
            if (!event.altKey && !event.shiftKey && !event.ctrlKey && !event.metaKey &&
                groupSelection.current === current.selectedRef.current && current.selectedRef.current.has(id)) {
                const ids = new Set([id]);
                current.selectedRef.current = ids;
                current.setSelected(ids);
            }
            groupSelection.current = null;
        },
        drag: (event, group) => {
            if (event.button !== 0 || resizing.current || !canInteract()) return;
            selectGroup(group);
            setMenu(null);
            // Use the existing node drag engine, including pointer cleanup and
            // batch children, but do not treat Option held during grouping as copy.
            startDrag({
                button: 0, pointerId: event.pointerId, clientX: event.clientX, clientY: event.clientY,
                altKey: false, shiftKey: false, metaKey: false, ctrlKey: false,
                preventDefault: () => event.preventDefault(),
                stopPropagation: () => event.stopPropagation(),
            }, group.nodes[0].id);
            groupSelection.current = latest.current.selectedRef.current;
            event.preventDefault();
        },
        resize: (event, group, direction) => {
            if (event.button !== 0 || resizing.current || !canInteract()) return;
            selectGroup(group);
            setMenu(null);
            resizing.current = beginCanvasGroupResize(event, group, direction,
                Math.max(.05, latest.current.viewport?.k || 1),
                frame => change(current => updateCanvasGroup(current, group.id, { frame })),
                () => { resizing.current = null; });
        },
        rename: (id, name) => { change(current => updateCanvasGroup(current, id, { name })); setEditing(null); },
        color: (id, color) => change(current => updateCanvasGroup(current, id, { color })),
        dissolve,
        remove: () => { change(current => assignCanvasGroup(current, actionIds(), null)); setMenu(null); },
        join: group => { change(current => assignCanvasGroup(current, actionIds(), group)); setMenu(null); },
    };
}

export function createCanvasGroupComponents(React, icons) {
    const h = React.createElement;
    const stop = event => event.stopPropagation();
    function IconButton({ title, icon, onClick }) {
        return h("button", { type: "button", title, "aria-label": title, onPointerDown: stop, onClick },
            h(icon, { size: 16, className: "size-4" }));
    }
    function NameEditor({ group, ui }) {
        const [value, setValue] = React.useState(group.name);
        const cancelled = React.useRef(false);
        return h("input", {
            autoFocus: true, value, maxLength: 80, "aria-label": "\u5206\u7ec4\u540d\u79f0",
            onFocus: event => event.target.select(), onPointerDown: stop,
            onChange: event => setValue(event.target.value),
            onBlur: () => { if (!cancelled.current) ui.rename(group.id, value); },
            onKeyDown: event => {
                event.stopPropagation();
                if (event.nativeEvent?.isComposing) return;
                if (event.key === "Enter") { event.preventDefault(); cancelled.current = true; ui.rename(group.id, value); }
                if (event.key === "Escape") { cancelled.current = true; ui.setEditing(null); }
            },
        });
    }
    function Layer({ ui, visibleIds, viewport, viewportSize }) {
        const groups = ui.groups.filter(group => {
            if (group.nodes.some(node => visibleIds.has(node.id))) return true;
            if (!viewport || !viewportSize) return false;
            const bounds = canvasGroupBounds(group.nodes);
            const left = bounds.x * viewport.k + viewport.x, top = bounds.y * viewport.k + viewport.y;
            return left < viewportSize.width && top < viewportSize.height &&
                left + bounds.width * viewport.k > 0 && top + bounds.height * viewport.k > 0;
        });
        return h("div", { className: "canvas-group-layer" },
            groups.map(group => {
                const bounds = canvasGroupBounds(group.nodes);
                const active = group.nodes.every(node => ui.selected.has(node.id));
                return h("div", {
                    key: group.id, className: `canvas-group-frame${active ? " is-selected" : ""}`,
                    "data-canvas-group-id": group.id,
                    style: { left: bounds.x, top: bounds.y, width: bounds.width, height: bounds.height, "--group-color": group.color },
                    onPointerDown: event => { if (!event.shiftKey) ui.drag(event, group); },
                    onDoubleClick: stop,
                    onContextMenu: event => { ui.selectGroup(group); ui.openMenu(event, group.id); },
                }, h("div", {
                    className: "canvas-group-heading", "data-canvas-no-zoom": true,
                    onPointerDown: event => ui.drag(event, group),
                    onDoubleClick: event => { event.stopPropagation(); ui.setEditing(group.id); },
                    onContextMenu: event => { ui.selectGroup(group); ui.openMenu(event, group.id); },
                },
                ui.editing === group.id ? h(NameEditor, { key: group.id, group, ui }) :
                    h("span", { className: "canvas-group-name", title: group.name }, group.name),
                h("span", { className: "canvas-group-count" }, group.nodes.length),
                active && h("div", { className: "canvas-group-tools", onDoubleClick: stop },
                    h(IconButton, { title: "\u4fee\u6539\u540d\u79f0", icon: icons.edit, onClick: () => ui.setEditing(group.id) }),
                    h("button", {
                        type: "button", className: "canvas-group-color-button", title: "\u5206\u7ec4\u989c\u8272", "aria-label": "\u5206\u7ec4\u989c\u8272",
                        onPointerDown: stop, onClick: event => ui.openMenu(event, group.id, null, true),
                    }, h("span", { style: { background: group.color } })),
                    h(IconButton, { title: "\u5206\u7ec4\u83dc\u5355", icon: icons.menu, onClick: event => ui.openMenu(event, group.id) }),
                )),
                ...["n", "ne", "e", "se", "s", "sw", "w", "nw"].map(direction => h("div", {
                    key: direction, className: `canvas-group-resize canvas-group-resize-${direction}`,
                    "data-canvas-no-zoom": true, title: "\u8c03\u6574\u7ec4\u6846\u5927\u5c0f",
                    onPointerDown: event => ui.resize(event, group, direction),
                    onDoubleClick: stop,
                })));
            }));
    }
    function Menu({ ui }) {
        const menu = ui.menu;
        const ref = React.useRef(null);
        React.useLayoutEffect(() => {
            if (!menu || !ref.current) return;
            const element = ref.current;
            element.style.left = `${Math.max(8, Math.min(menu.x, window.innerWidth - element.offsetWidth - 8))}px`;
            element.style.top = `${Math.max(8, Math.min(menu.y, window.innerHeight - element.offsetHeight - 8))}px`;
        }, [menu, ui.groups.length]);
        React.useEffect(() => {
            if (menu) ref.current?.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
        }, [menu]);
        if (!menu) return null;
        const group = ui.groups.find(item => item.id === menu.groupId);
        const selected = new Set(menu.nodeIds);
        const action = (label, handler, disabled = false) => h("button", {
            key: label, type: "button", role: "menuitem", disabled,
            onClick: () => { handler(); ui.setMenu(null); },
        }, label);
        return h("div", {
            ref, className: "canvas-group-menu", role: "menu", "aria-label": "\u5206\u7ec4",
            "data-canvas-group-menu": true, "data-canvas-no-zoom": true,
            style: { left: menu.x, top: menu.y },
            onPointerDown: stop, onContextMenu: event => event.preventDefault(),
            onKeyDown: event => {
                event.stopPropagation();
                if (event.key === "Escape") { event.preventDefault(); ui.setMenu(null); }
                if (event.key === "ArrowDown" || event.key === "ArrowUp") {
                    event.preventDefault();
                    const controls = [...ref.current.querySelectorAll("button:not(:disabled),input")];
                    const index = controls.indexOf(document.activeElement);
                    controls[(index + (event.key === "ArrowDown" ? 1 : controls.length - 1)) % controls.length]?.focus();
                }
            },
        }, group ? [
            !menu.palette && action("\u4fee\u6539\u540d\u79f0", () => ui.setEditing(group.id)),
            h("div", { key: "colors", className: "canvas-group-swatches" },
                GROUP_COLORS.map(color => h("button", {
                    key: color, type: "button", title: color, "aria-label": color, "aria-pressed": group.color === color,
                    style: { background: color }, onClick: () => ui.color(group.id, color),
                })),
                h("input", { type: "color", value: group.color, title: "\u81ea\u5b9a\u4e49\u989c\u8272", "aria-label": "\u81ea\u5b9a\u4e49\u989c\u8272",
                    onChange: event => ui.color(group.id, event.target.value) })),
            !menu.palette && action("\u5220\u9664\u7ec4\u6846\uff08\u4fdd\u7559\u5185\u5bb9\uff09", () => ui.dissolve(group.id)),
        ] : [
            action("\u521b\u5efa\u5206\u7ec4", ui.create, selected.size < 2),
            action("\u79fb\u51fa\u5206\u7ec4", ui.remove, !ui.groups.some(item => item.nodes.some(node => selected.has(node.id)))),
            ...ui.groups.map(item => action(`\u52a0\u5165\u5206\u7ec4 \u00b7 ${item.name}`, () => ui.join(item))),
        ]);
    }
    return { Layer, Menu };
}
