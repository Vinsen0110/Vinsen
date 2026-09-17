import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const bundle = await readFile(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");
const indexHtml = await readFile(new URL("../index.html", import.meta.url), "utf8");

test("settings site tabs only change the site being edited", () => {
    assert.match(bundle, /\[configSiteId,setConfigSiteId\]=c\.useState\(null\)/);
    assert.match(
        bundle,
        /k=\$=>\{setEditingKeyId\(null\),setKeyMenuOpen\(!1\),setConfigSiteId\(\$\)\}/,
    );
    assert.doesNotMatch(
        bundle,
        /k=\$=>\{setEditingKeyId\(null\),setKeyMenuOpen\(!1\);const j=buildSiteConfigPatch/,
    );
    assert.match(
        bundle,
        /j!==void 0&&v===A\.activeSiteId\?\{apiKey:j\}:\{\}/,
        "editing an inactive site must not replace the active top-level API key",
    );
});

test("API key edits read the latest site config before readiness and close", () => {
    assert.match(
        bundle,
        /updateSite=\(\$,j\)=>\{const A=Xr\.getState\?\.\(\)\.config\|\|a,T=A\.channels\.map\(N=>N\.id===v\?\{\.\.\.N,\.\.\.\$\}:N\)/,
    );
    assert.match(
        bundle,
        /C=\(\)=>\{setEditingKeyId\(null\),setKeyMenuOpen\(!1\);const A=Xr\.getState\?\.\(\)\.config\|\|a,I=A\.channels\.find\(N=>N\.id===v\)\|\|g,K=normalizeSiteApiKeys\(I\)\.apiKey;f\(!1\),K\.trim\(\)&&/,
    );
    assert.doesNotMatch(
        bundle,
        /C=\(\)=>\{setEditingKeyId\(null\),setKeyMenuOpen\(!1\),f\(!1\),a\.apiKey\.trim\(\)&&/,
    );
});

test("the bottom-left control exposes explicit site choices", () => {
    assert.match(bundle, /className:"canvas-site-menu",role:"menu"/);
    assert.match(bundle, /role:"menuitemradio","aria-checked":f\.id===activeSite\?\.id/);
    assert.match(bundle, /onClick:\(\)=>selectSite\(f\.id\)/);
    assert.match(
        bundle,
        /selectSite=f=>\{const m=buildSiteConfigPatch\(siteConfig,f\);m&&\(patchSiteConfig\(m\),setSiteMenuOpen\(!1\)\)\}/,
    );
    assert.doesNotMatch(bundle, /switchSite=\(\)=>/);
    assert.match(indexHtml, /\.canvas-site-menu-item\.is-active/);
});

test("the current canvas site name has stronger visual emphasis", () => {
    assert.match(
        bundle,
        /className:"canvas-site-current-label","aria-live":"polite",children:\["\\u5F53\\u524D\\uFF1A",y\.jsx\("strong"/,
    );
    assert.match(indexHtml, /\.canvas-site-current-label strong \{[^}]*font-weight: 800/s);
    assert.match(indexHtml, /index-B2KJ37fm\.js\?v=[^"\s]+/);
});

test("the lower-left toolbar only shows site, minimap, and zoom controls", () => {
    assert.doesNotMatch(
        bundle,
        /y\.jsx\(Eo,\{title:t\?"\\u5173\\u95ED\\u5C0F\\u5730\\u56FE":"\\u6253\\u5F00\\u5C0F\\u5730\\u56FE"/,
        "the minimap button must not create a tooltip",
    );
    assert.doesNotMatch(
        bundle,
        /y\.jsx\(Eo,\{title:"\\u5FEB\\u6377\\u952E",children:y\.jsx\(Bt/,
        "the lower-left shortcut control must not be rendered",
    );
    assert.doesNotMatch(
        bundle,
        /y\.jsx\(Lr,\{title:"\\u5FEB\\u6377\\u952E"/,
        "the unreachable shortcut modal must not remain in the canvas bundle",
    );
    assert.match(
        indexHtml,
        /\.canvas-site-switch-button ~ button\[aria-label="快捷键"\][^}]*display: none !important;/s,
    );
    assert.match(indexHtml, /body:has\(\.canvas-site-switch-button:hover\) \.ant-tooltip,[^}]*display: none !important;/s);
    assert.match(indexHtml, /body:has\(\.canvas-site-menu\) \.ant-tooltip \{[^}]*display: none !important;/s);
});

test("the redundant disabled redo icon is hidden from the left toolbar", () => {
    assert.doesNotMatch(bundle, /id:"tool-redo"/);
    assert.match(
        indexHtml,
        /button\[aria-label="重做 Ctrl \/ Cmd \+ Shift \+ Z"\] \{[^}]*display: none !important;/s,
    );
});

test("generation defaults identify the active site and use a readable select state", () => {
    assert.match(
        bundle,
        /label:"\\u9ED8\\u8BA4\\u751F\\u56FE\\u6A21\\u578B\\uFF08"\+\(activeSiteChannel\(a\)\?\.name/,
    );
    assert.match(bundle, /className:"generation-settings-sections"/);
    assert.match(bundle, /children:"\\u751F\\u56FE\\u6A21\\u578B\\u8BBE\\u7F6E"/);
    assert.match(bundle, /children:"\\u6587\\u672C\\u6A21\\u578B\\u8BBE\\u7F6E"/);
    assert.match(bundle, /popupClassName:"settings-select-dropdown"/);
    assert.match(
        indexHtml,
        /\.settings-select-dropdown \.ant-select-item-option-selected[^}]*background: #eaf2ff !important;[^}]*color: #175cd3 !important;/s,
    );
    assert.match(
        indexHtml,
        /\.settings-select\.ant-select-open \.ant-select-selector \.ant-select-selection-item[^}]*color: #1f2937 !important;[^}]*-webkit-text-fill-color: #1f2937 !important;[^}]*opacity: 1 !important;/s,
    );
    assert.match(
        indexHtml,
        /\.dark \.settings-select-dropdown[^}]*background: #1c1917 !important;[^}]*color: #e7e5e4 !important;/s,
    );
    assert.match(bundle, /className:"settings-model-select !h-11/);
    assert.match(
        indexHtml,
        /\.settings-model-select\[data-state="open"\] \.canvas-model-picker-text[^}]*color: #1f2937 !important;[^}]*-webkit-text-fill-color: #1f2937 !important;/s,
    );
    assert.match(
        indexHtml,
        /\.dark \.settings-model-select\[data-state="open"\] \.canvas-model-picker-text[^}]*color: #f5f5f4 !important;[^}]*-webkit-text-fill-color: #f5f5f4 !important;/s,
    );
});
