export const PANEL_RESIZE_EVENT = "canvas-generation-panel-resize";

const validSize = value => typeof value === "number" && Number.isFinite(value) && value > 0;

export function bindCanvasPanelSize(panel, metadata = {}, onResizePatch) {
    const saved = {};
    for (const axis of ["width", "height"]) {
        const suffix = axis === "width" ? "Width" : "Height";
        const key = `generationPanel${suffix}`;
        const value = metadata[key];
        const property = `--canvas-panel-resize-${axis}`;
        if (validSize(value)) {
            saved[key] = value;
            panel.style.setProperty(property, `${value}px`);
            panel.dataset[`canvasResize${suffix}Ready`] = "true";
        } else {
            panel.style.removeProperty(property);
            delete panel.dataset[`canvasResize${suffix}Ready`];
        }
    }
    if (saved.generationPanelWidth) panel.dataset.canvasWidthLocked = "true";
    else delete panel.dataset.canvasWidthLocked;
    if (Object.keys(saved).length) panel.dataset.canvasResizeReady = "true";
    else delete panel.dataset.canvasResizeReady;

    const onResize = event => {
        if (event.target !== panel) return;
        const patch = {};
        for (const axis of ["width", "height"]) {
            const key = axis === "width" ? "generationPanelWidth" : "generationPanelHeight";
            const value = event.detail?.[axis];
            if (validSize(value) && value !== saved[key]) patch[key] = value;
        }
        if (!Object.keys(patch).length) return;
        Object.assign(saved, patch);
        onResizePatch(patch);
    };
    panel.addEventListener(PANEL_RESIZE_EVENT, onResize);
    return () => panel.removeEventListener(PANEL_RESIZE_EVENT, onResize);
}
