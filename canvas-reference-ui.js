export function mentionMenuLayout(anchor, bounds, measured, preferredSide) {
    const gap = 6;
    const width = Math.min(256, Math.max(0, bounds.right - bounds.left));
    const below = Math.max(0, bounds.bottom - anchor.bottom - gap);
    const above = Math.max(0, anchor.top - bounds.top - gap);
    const desired = Math.min(240, measured);
    let side = preferredSide;
    if (!side || (side === "bottom" ? below : above) < desired) {
        side = below >= desired || below >= above ? "bottom" : "top";
    }
    const height = Math.min(desired, side === "bottom" ? below : above);
    const left = Math.max(bounds.left, Math.min(anchor.left, bounds.right - width));
    const top = Math.max(bounds.top, Math.min(
        side === "bottom" ? anchor.bottom + gap : anchor.top - gap - height,
        bounds.bottom - height,
    ));
    return { left, top, width, height, side };
}

export function keepMentionVisible(list, index) {
    const item = list?.children[index];
    if (!item) return;
    const top = item.offsetTop;
    const bottom = top + item.offsetHeight;
    if (top < list.scrollTop) list.scrollTop = top;
    else if (bottom > list.scrollTop + list.clientHeight) {
        list.scrollTop = bottom - list.clientHeight;
    }
}

export function observeMentionMenu(anchor, menu, getCaretRect) {
    let frame = null;
    let side;
    const list = menu.querySelector('[role="listbox"]');
    const update = () => {
        frame = null;
        if (!anchor.isConnected || !menu.isConnected) return;
        const editor = anchor.getBoundingClientRect();
        const caret = getCaretRect() || editor;
        // A scrolled-out caret must not anchor the menu outside the editor.
        const rect = {
            left: Math.max(editor.left, Math.min(caret.left, editor.right)),
            top: Math.max(editor.top, Math.min(caret.top, editor.bottom)),
            bottom: Math.max(editor.top, Math.min(caret.bottom, editor.bottom)),
        };
        const modal = anchor.closest(".ant-modal-content")?.getBoundingClientRect();
        const bounds = {
            left: Math.max(8, modal?.left + 8 || 8),
            top: Math.max(8, modal?.top + 8 || 8),
            right: Math.min(window.innerWidth - 8, modal?.right - 8 || window.innerWidth - 8),
            bottom: Math.min(window.innerHeight - 8, modal?.bottom - 8 || window.innerHeight - 8),
        };
        const header = menu.firstElementChild;
        const layout = mentionMenuLayout(rect, bounds, (header?.offsetHeight || 0) + (list?.scrollHeight || 0) + 2, side);
        side = layout.side;
        for (const key of ["left", "top", "width", "height"]) {
            const value = layout[key] + "px";
            if (menu.style[key] !== value) menu.style[key] = value;
        }
        menu.style.visibility = "visible";
    };
    const schedule = event => {
        if (event?.target && menu.contains(event.target)) return;
        if (frame === null) frame = requestAnimationFrame(update);
    };
    const resize = new ResizeObserver(schedule);
    resize.observe(anchor);
    const panel = anchor.closest(".canvas-generation-panel");
    if (panel) resize.observe(panel);
    const movement = new MutationObserver(schedule);
    for (let element = anchor.parentElement; element; element = element.parentElement) {
        movement.observe(element, { attributes: true, attributeFilter: ["style", "class"] });
        if (element.classList.contains("laowu-canvas-viewport")) break;
    }
    movement.observe(anchor, { childList: true, subtree: true, characterData: true });
    window.addEventListener("resize", schedule);
    window.addEventListener("scroll", schedule, true);
    document.addEventListener("selectionchange", schedule);
    update();
    return () => {
        if (frame !== null) cancelAnimationFrame(frame);
        resize.disconnect();
        movement.disconnect();
        window.removeEventListener("resize", schedule);
        window.removeEventListener("scroll", schedule, true);
        document.removeEventListener("selectionchange", schedule);
    };
}

export function remapImageMentions(prompt, before, after) {
    if (typeof prompt !== "string") return prompt;
    const next = new Map(after.map((id, index) => [id, index + 1]));
    return prompt.replace(/(@?)图片(\d+)/g, (token, prefix, number) => {
        const id = before[Number(number) - 1];
        if (id === undefined) return token;
        const index = next.get(id);
        return index === undefined ? "" : `${prefix}图片${index}`;
    });
}

export function reorderReferenceEdges(connections, targetId, before, after) {
    if (before.length === after.length && before.every((id, index) => id === after[index])) return connections;
    const allowed = new Set(before);
    if (new Set(after).size !== after.length || after.some(id => !allowed.has(id))) return connections;
    const edges = new Map(connections.filter(edge => edge.toNodeId === targetId && allowed.has(edge.fromNodeId))
        .map(edge => [edge.fromNodeId, edge]));
    if (after.some(id => !edges.has(id))) return connections;
    const ordered = after.map(id => edges.get(id));
    let index = 0;
    return connections.flatMap(edge => {
        if (edge.toNodeId !== targetId || !allowed.has(edge.fromNodeId)) return [edge];
        return index < ordered.length ? [ordered[index++]] : [];
    });
}

export function installCanvasFlowActivity() {
    let viewport = null;
    let timer = null;
    const resume = () => {
        clearTimeout(timer);
        timer = null;
        viewport?.removeAttribute("data-canvas-interacting");
        viewport = null;
    };
    const pause = event => {
        const target = event.target instanceof Element ? event.target : null;
        if (target?.closest("[data-canvas-no-zoom],input,textarea,[contenteditable=true]")) return;
        const canvas = target?.closest(".laowu-canvas-viewport");
        if (!canvas) return;
        if (viewport !== canvas) resume();
        viewport = canvas;
        if (!viewport.hasAttribute("data-canvas-interacting")) {
            viewport.setAttribute("data-canvas-interacting", "true");
        }
        clearTimeout(timer);
        timer = setTimeout(resume, 160);
    };
    const move = event => { if (event.buttons) pause(event); };
    const visibility = () => {
        document.documentElement.dataset.canvasPageHidden = String(document.hidden);
        if (document.hidden) resume();
    };
    visibility();
    document.addEventListener("visibilitychange", visibility);
    document.addEventListener("pointermove", move, { passive: true });
    document.addEventListener("wheel", pause, { passive: true });
    window.addEventListener("pointerup", resume);
    window.addEventListener("pointercancel", resume);
    window.addEventListener("blur", resume);
    return () => {
        resume();
        document.removeEventListener("visibilitychange", visibility);
        document.removeEventListener("pointermove", move);
        document.removeEventListener("wheel", pause);
        window.removeEventListener("pointerup", resume);
        window.removeEventListener("pointercancel", resume);
        window.removeEventListener("blur", resume);
    };
}

let flowObserver;
let observedFlowLines = 0;
export function observeFlowLine(line) {
    if (!line || typeof IntersectionObserver !== "function") return;
    if (!flowObserver) flowObserver = new IntersectionObserver(entries => {
        for (const entry of entries) {
            entry.target.dataset.canvasFlowVisible = String(entry.isIntersecting);
        }
    });
    line.dataset.canvasFlowVisible = "false";
    flowObserver.observe(line);
    observedFlowLines++;
    return () => {
        flowObserver.unobserve(line);
        if (--observedFlowLines === 0) {
            flowObserver.disconnect();
            flowObserver = undefined;
        }
    };
}
