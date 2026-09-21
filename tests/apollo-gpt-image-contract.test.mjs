import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

const bundle = fs.readFileSync(new URL("../assets/index-B2KJ37fm.js", import.meta.url), "utf8");

test("Apollo removes GPT Image 2 from the catalog but retains GPT Image 2.5", () => {
    assert.match(bundle, /APOLLO_IMAGE_MODELS=\["nano-banana-pro","gpt-image-2\.5"\]/);
    assert.match(bundle, /APOLLO_SITE_MODELS=\["nano-banana-pro-2k","nano-banana-pro-4k","nano-banana-pro","gpt-image-2\.5"/);
    assert.match(bundle, /"gpt-image-2\.5","gemini-3\.8-flash"\]/);
});

test("Apollo GPT Image 2 keeps request model and resolution mapping isolated", () => {
    assert.match(bundle, /function isApolloGptImageModel\(e\)/);
    assert.match(bundle, /function apolloGptImageRequestModel\(e\)\{const t=pr\(e\?\.model\|\|e\?\.imageModel\|\|"gpt-image-2"\)/);
    assert.match(bundle, /model:apolloGptImageRequestModel\(e\),prompt:Zq\(e,t\),size:apolloGptImageSize\(e\),quality:tudouQuality\(e\)/);
    assert.match(bundle, /t==="auto"\?816:t==="1k"\?1024:t==="2k"\?2048:2880/);
    assert.match(bundle, /Math\.sqrt\(i\*a\).*Math\.sqrt\(i\/a\)/);
    assert.match(bundle, /for\(;m\*h>TMe;\).*for\(;m\*h<MMe;\)/);
});

test("Apollo GPT Image 2 uses the RH quality and extended ratio controls", () => {
    assert.match(bundle, /isGptImageModel=\["gpt-image-2","gpt-image-2\.5"\]\.includes\(pr\(e\.model\)\)&&\(isApilioSite\(e\)\|\|isTudouSite\(e\)\|\|isRunningHubSite\(e\)/);
    assert.match(bundle, /ratioPresets=orderRatioPresets\(isRunningHub25\?runningHub25RatioOptions\(e,\[\.\.\.pg,\.\.\.GPT_IMAGE_EXTRA_RATIO_PRESETS\]\):isGptImageModel\?\[\.\.\.pg,\.\.\.GPT_IMAGE_EXTRA_RATIO_PRESETS\]/);
    assert.match(bundle, /gptQualityOptions=GPT_IMAGE_QUALITY_OPTIONS/);
    assert.match(bundle, /children:gptQualityOptions\.map\(k=>y\.jsx\(GL/);
});

test("Apollo GPT Image 2 shows quality in the compact canvas toolbar", () => {
    assert.match(bundle, /isGptImageConfig=\["gpt-image-2","gpt-image-2-vip","gpt-image-2\.5"\]\.includes\(pr\(a\.imageModel\|\|a\.model\)\)&&\["apilio","tudou","runninghub","grsai","apimart"\]\.includes/);
    assert.match(bundle, /u=Array\.isArray\(t\.channels\)\?bd\(t,t\.model\):t/);
    assert.match(bundle, /isApilioSite\(u\)\|\|isTudouSite\(u\)\|\|isRunningHubSite\(u\)\|\|isGrsaiSite\(u\)\|\|isApiMartSite\(u\)/);
    assert.match(bundle, /y\.jsx\(n\$,\{value:w,items:GPT_IMAGE_QUALITY_OPTIONS/);
});

test("Apollo GPT Image 2 always submits async generations and edits", () => {
    assert.match(bundle, /submitImageRequest\(e,n\.length\?"\/images\/edits":"\/images\/generations",i,n\.length\?wA\(e\):wA\(e,"application\/json"\),o,!0,a,!0\)/);
    assert.match(bundle, /if\(j\|\|!imageAsyncUnsupported\(h\)\)throw h/);
    assert.match(bundle, /`\$\{isTudouSite\(e\)\?"\/tasks":"\/images\/tasks"\}\//);
});

test("Apollo GPT Image 2.5 edits infer source ratio when size is auto", () => {
    assert.match(bundle, /async function withApolloGptReferenceSize\(e,t\)/);
    assert.match(bundle, /String\(e\?\.size\|\|""\)\.trim\(\)\.toLowerCase\(\)!=="auto"/);
    assert.match(bundle, /Math\.max\(o,a\)\/Math\.min\(o,a\)>3/);
    assert.match(bundle, /size:`\$\{o\}:\$\{a\}`/);
    assert.match(bundle, /const h0=await withApolloGptReferenceSize\(a,n\);const d=await Promise\.all/);
});

test("Apollo GPT Image 2 edits send the documented multipart fields", () => {
    assert.match(bundle, /r\.set\("model",apolloGptImageRequestModel\(e\)\),r\.set\("prompt",Zq\(e,t\)\),r\.set\("size",apolloGptImageSize\(e\)\),r\.set\("quality",tudouQuality\(e\)\)/);
    assert.match(bundle, /isApolloGptImageModel\(a\.model\).*nM\(\{\.\.\.h,dataUrl:await vh\(h\)\}\)/);
});

test("Apollo keeps billing metadata per key and never changes the request model", () => {
    assert.match(bundle, /from"\.\.\/apollo-billing\.js"/);
    assert.match(bundle, /billingGroup:normalizeApolloBillingGroup\(n\?\.billingGroup\)/);
    assert.match(bundle, /detectedBillingGroup:normalizeApolloDetectedBillingGroup\(n\?\.detectedBillingGroup\)/);
    assert.match(bundle, /apolloBillingKey:activeSiteApiKey\(n\)/);
    assert.match(bundle, /i&&n==="gpt-image-2"\?apolloGptImagePrice/);
    assert.match(bundle, /children:"密钥计费组"/);
    assert.match(bundle, /children:"\u91CD\u65B0\u68C0\u6D4B"/);
    assert.match(bundle, /children:"\u5F85\u786E\u8BA4"/);
    assert.match(bundle, /\(\?:\u666E\u901A\|default\).*\(\?:\u4F18\u8D28\|premium\|official\).*\?"label"/);
    assert.match(bundle, /function assertApolloNanoBilling\(e\)/);
    assert.match(bundle, /effectiveApolloBillingGroup\(e\.apolloBillingKey\)===APOLLO_BILLING_GROUP_DEFAULT/);
    assert.match(bundle, /function apolloGptImageRequestModel\(/, "resolution routing must stay outside billing metadata");
    assert.doesNotMatch(bundle, /model:"gpt-image-2-official-mix"/);
});

test("Apollo exposes GPT Image 2.5 with variant mapping and billing-group pricing", () => {
    assert.match(bundle, /isApilioGpt25/);
    assert.match(bundle, /function apolloGptImageRequestModel\(e\)\{const t=pr\(e\?\.model\|\|e\?\.imageModel\|\|"gpt-image-2"\);return t==="gpt-image-2\.5"/);
    assert.match(bundle, /i&&n==="gpt-image-2\.5"\?apolloGptImage25Price/);
    assert.match(bundle, /isApilioGpt25\?y\.jsx\(n\$,\{value:apiMart25VariantValue/);
});

test("API settings remove retired GPT2 pricing while preserving shared key controls", () => {
    const settings = bundle.slice(bundle.indexOf("function t$e("), bundle.indexOf('function nL('));
    assert.ok(settings.length > 0);
    assert.doesNotMatch(settings, /GPT Image 2 计费组|billingPrice|apolloGptImagePrice\(|api-billing-price/);
    assert.match(settings, /children:"密钥计费组"/);
    assert.match(settings, /onChange:setActiveKeyBillingGroup/);
    assert.match(settings, /onClick:\(\)=>detectActiveKeyBilling\(!0\)/);
    assert.match(settings, /onClick:addSiteKey/);
    assert.match(settings, /onClick:deleteSiteKey/);
    assert.match(settings, /value:a\.totalBalanceUserId/);
    assert.match(settings, /value:a\.totalBalanceToken/);
});
