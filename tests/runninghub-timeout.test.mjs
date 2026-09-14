import assert from "node:assert/strict";
import test from "node:test";
import { runRunningHubImageGeneration } from "../runninghub-api.js";

const config = {
    apiKey: "mock-key",
    provider: "runninghub",
    model: "gpt-image-2.5",
    runningHubGpt25Mode: "fixed",
    runningHubGpt25Variant: "sunburst",
    quality: "4k",
    size: "16:9",
};

function clockedTask({ completesAt = Infinity, controller, cancelAfterQueries } = {}) {
    let elapsed = 0;
    const requests = [];
    const sleeps = [];
    const taskId = "mock-rh-task";
    const resultUrl = "https://example.test/rh-result.png";
    const options = {
        signal: controller?.signal,
        now: () => elapsed,
        sleep: async ms => {
            sleeps.push(ms);
            elapsed += ms;
            const queryCount = requests.filter(request => request.url.endsWith("/query")).length;
            if (controller && queryCount === cancelAfterQueries) {
                controller.abort(new DOMException("Cancelled by user", "AbortError"));
            }
        },
        fetchImpl: async (url, init) => {
            requests.push({ url, at: elapsed, body: JSON.parse(init.body), signal: init.signal });
            const query = url.endsWith("/query");
            return new Response(JSON.stringify(query && elapsed >= completesAt ? {
                taskId,
                status: "SUCCESS",
                results: [{ url: resultUrl }],
            } : {
                taskId,
                status: "RUNNING",
                results: null,
            }), {
                status: 200,
                headers: { "Content-Type": "application/json" },
            });
        },
    };
    return {
        options, requests, sleeps, resultUrl, taskId,
        elapsed: () => elapsed,
        queries: () => requests.filter(request => request.url.endsWith("/query")),
        submissions: () => requests.filter(request => !request.url.endsWith("/query")),
    };
}

test("RH receives a task completed at 7m22s without resubmitting after the old five-minute boundary", async () => {
    const task = clockedTask({ completesAt: 7 * 60_000 + 22_000 });
    const result = await runRunningHubImageGeneration(config, "draw", [], task.options);
    assert.deepEqual(result, [task.resultUrl]);
    assert.equal(task.submissions().length, 1);
    assert.equal(task.elapsed(), 443_000, "the first scheduled query after completion returns the result");
    assert.ok(task.queries().some(request => request.at > 300_000));
    assert.ok(task.queries().every(request => request.body.taskId === task.taskId));
});

test("RH stops starting queries after ten minutes and submits only one billable task", async () => {
    const task = clockedTask();
    await assert.rejects(
        runRunningHubImageGeneration(config, "draw", [], task.options),
        /RH 图像生成超时/,
    );
    assert.equal(task.submissions().length, 1);
    assert.equal(task.queries().length, 200);
    assert.equal(task.queries()[0].at, 2_000);
    assert.equal(task.queries().at(-1).at, 599_000);
    assert.ok(task.queries().every(request => request.at < 600_000));
    assert.equal(task.elapsed(), 602_000, "the normal 3s sleep may cross the limit, but no extra query starts");
    assert.ok(task.queries().every(request => request.body.taskId === task.taskId));
});

test("RH keeps its first 2s delay and normal 3s query interval", async () => {
    const task = clockedTask({ completesAt: 8_000 });
    await runRunningHubImageGeneration(config, "draw", [], task.options);
    assert.deepEqual(task.sleeps, [2_000, 3_000, 3_000]);
    assert.deepEqual(task.queries().map(request => request.at), [2_000, 5_000, 8_000]);
    assert.equal(task.submissions().length, 1);
});

test("user cancellation still stops RH polling immediately without resubmitting", async () => {
    const controller = new AbortController();
    const task = clockedTask({ controller, cancelAfterQueries: 2 });
    await assert.rejects(
        runRunningHubImageGeneration(config, "draw", [], task.options),
        error => error === controller.signal.reason && error.name === "AbortError",
    );
    assert.equal(task.queries().length, 2);
    assert.equal(task.submissions().length, 1);
    assert.ok(task.requests.every(request => request.signal === controller.signal));
});

test("an explicit RH task timeout still overrides the longer default", async () => {
    const task = clockedTask();
    await assert.rejects(
        runRunningHubImageGeneration(config, "draw", [], { ...task.options, timeoutMs: 7_000 }),
        /RH 图像生成超时/,
    );
    assert.deepEqual(task.queries().map(request => request.at), [2_000, 5_000]);
    assert.equal(task.submissions().length, 1);
});
