#!/usr/bin/env node

// Offline regression suite for the build-request retry policy
// (src/apiService.ts: sendBuildRequest).
//
// Verifies: only transient failures retry (408/429/502/503/504 + network
// exceptions), permanent failures fail fast without sleeping, Drive-save
// errors never re-POST the build, and sendBuildRequest itself shows no
// blocking ui.alert (the caller owns user-facing errors).

const {
  collectTypeScriptFiles,
  createGasContext,
  loadFiles,
  FakeSheet,
  createFakeSpreadsheet,
} = require("./gasSimHarness.cjs");

const ignoredDirs = new Set(["scripts"]);

function makeResponse(code, body, contentType) {
  return {
    getResponseCode() {
      return code;
    },
    getContentText() {
      return body;
    },
    getHeaders() {
      return contentType ? { "Content-Type": contentType } : {};
    },
    getBlob() {
      return {
        setName(name) {
          this.name = name;
          return this;
        },
      };
    },
  };
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(`check-api-retry FAILED: ${message}`);
  }
  console.log(`  ok: ${message}`);
}

function createTest({ fetchScript }) {
  const { spreadsheet, toasts } = createFakeSpreadsheet({
    Categories: new FakeSheet("Categories", [["English"]]),
  });
  const state = {
    fetches: 0,
    sleeps: [],
    alerts: 0,
    driveSaves: 0,
    toasts,
  };

  const context = createGasContext();
  context.SpreadsheetApp = {
    getActiveSpreadsheet() {
      return spreadsheet;
    },
    getUi() {
      return {
        Button: { NO: "NO", YES: "YES" },
        ButtonSet: { OK: "OK", YES_NO: "YES_NO" },
        alert() {
          state.alerts++;
          return "OK";
        },
      };
    },
  };
  context.Utilities.sleep = (ms) => {
    state.sleeps.push(ms);
  };
  context.UrlFetchApp = {
    fetch() {
      state.fetches++;
      const step = fetchScript.shift();
      if (step instanceof Error) {
        throw step;
      }
      return step;
    },
  };
  loadFiles(context, collectTypeScriptFiles(ignoredDirs));

  // Stub AFTER loadFiles: the real declarations overwrite any pre-load stubs.
  context.validateBuildRequest = () => {};
  context.saveComapeocatToDrive = (blob) => {
    state.driveSaves++;
    state.savedBlob = blob;
    return "drive-url";
  };
  context.closeProcessingModalDialog = () => {};

  const request = { metadata: { name: "t", version: "1" } };

  return { context, state, run: () => context.sendBuildRequest(request) };
}

const BINARY = "application/octet-stream";
const OK = makeResponse(200, "", BINARY);

function runCases() {
  // 1) Immediate success: one fetch, no backoff sleeps.
  let t = createTest({ fetchScript: [OK] });
  let result = t.run();
  assert(result === "drive-url", "200 binary returns the saved Drive URL");
  assert(t.state.fetches === 1, `single fetch on success (got ${t.state.fetches})`);
  assert(t.state.sleeps.length === 0, "no backoff sleep on success");
  assert(t.state.driveSaves === 1, "blob saved to Drive once");

  // 2) Transient 503s retry with exponential backoff, then succeed.
  t = createTest({ fetchScript: [makeResponse(503, "unavailable"), makeResponse(503, "unavailable"), OK] });
  result = t.run();
  assert(result === "drive-url", "503/503/200 eventually succeeds");
  assert(t.state.fetches === 3, `three fetches for two 503s (got ${t.state.fetches})`);
  assert(JSON.stringify(t.state.sleeps) === "[1000,2000]", `backoff 1000ms then 2000ms (got ${JSON.stringify(t.state.sleeps)})`);
  assert(t.state.toasts.filter((toast) => toast.title === "Retrying build request").length === 2, "retry notice via non-blocking toast");
  assert(t.state.alerts === 0, "no blocking ui.alert on the retry path");

  // 3) Permanent failures fail fast: 400/401/422/500, one fetch, no sleep.
  for (const code of [400, 401, 422, 500]) {
    t = createTest({ fetchScript: [makeResponse(code, JSON.stringify({ message: `bad request ${code}` }))] });
    let threw = null;
    try {
      t.run();
    } catch (error) {
      threw = error;
    }
    assert(threw && threw.message.includes(`rejected the request (${code})`), `${code} fails fast with "rejected" message`);
    assert(t.state.fetches === 1, `${code} is not retried (fetches: ${t.state.fetches})`);
    assert(t.state.sleeps.length === 0, `${code} does not sleep`);
  }

  // 4) Retryable-but-persistent 429 exhausts attempts with backoff.
  t = createTest({ fetchScript: [makeResponse(429, "slow down"), makeResponse(429, "slow down"), makeResponse(429, "slow down")] });
  let threw = null;
  try {
    t.run();
  } catch (error) {
    threw = error;
  }
  assert(threw && /after 3 attempts/.test(threw.message), `exhausted retries throw "after 3 attempts" (got: ${threw && threw.message})`);
  assert(JSON.stringify(t.state.sleeps) === "[1000,2000]", "429 exhaustion still backs off between attempts");
  assert(t.state.alerts === 0, "no blocking ui.alert when retries are exhausted");

  // 5) Network exceptions retry (transient).
  t = createTest({ fetchScript: [new Error("Address unavailable: http://x"), OK] });
  result = t.run();
  assert(result === "drive-url", "transient network exception then success");

  // 6) UrlFetchApp quota errors abort immediately.
  t = createTest({ fetchScript: [new Error("Service invoked too many times for one day: urlfetch.")] });
  threw = null;
  try {
    t.run();
  } catch (error) {
    threw = error;
  }
  assert(threw && /Could not reach build server/.test(threw.message), "quota fetch error aborts without retrying");
  assert(t.state.fetches === 1, `quota error is not retried (fetches: ${t.state.fetches})`);

  // 7) 200 with a non-binary body is a protocol mismatch: fail fast.
  t = createTest({ fetchScript: [makeResponse(200, JSON.stringify({ message: "bad gateway payload" }), "application/json")] });
  threw = null;
  try {
    t.run();
  } catch (error) {
    threw = error;
  }
  assert(threw && threw.message.includes("bad gateway payload"), `200 non-binary fails fast (got: ${threw && threw.message})`);
  assert(t.state.fetches === 1, "200 non-binary is not retried");

  // 8) Drive-save failure propagates and must NOT re-POST the build.
  t = createTest({ fetchScript: [OK] });
  t.context.saveComapeocatToDrive = () => {
    t.state.driveSaves++;
    throw new Error("Drive: storage quota exceeded");
  };
  threw = null;
  try {
    t.run();
  } catch (error) {
    threw = error;
  }
  assert(threw && /Drive: storage quota exceeded/.test(threw.message), "Drive-save error propagates to the caller");
  assert(t.state.fetches === 1, `Drive-save error does not re-POST the build (fetches: ${t.state.fetches})`);

  // 9) sendBuildRequest never blocks on ui.alert; the caller owns dialogs.
  t = createTest({ fetchScript: [makeResponse(500, "boom")] });
  try {
    t.run();
  } catch {
    // expected
  }
  assert(t.state.alerts === 0, `no ui.alert from sendBuildRequest (got ${t.state.alerts})`);
}

function main() {
  console.log("check-api-retry:");
  runCases();
  console.log("check-api-retry PASSED");
}

main();
