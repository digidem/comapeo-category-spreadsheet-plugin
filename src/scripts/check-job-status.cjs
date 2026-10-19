#!/usr/bin/env node

// Offline regression suite for the job status store (src/jobStatus.ts,
// issue #57). Loads the project sources through the shared gasSimHarness
// and exercises the set/progress/get/clear lifecycle against stateful fake
// PropertiesService + CacheService stores, plus the cache-eviction
// fallback, terminal-state locking, and the sidebar launcher.

const { collectTypeScriptFiles, createGasContext, loadFiles } = require("./gasSimHarness.cjs");

const ignoredDirs = new Set(["scripts", "test"]);

function assert(condition, message) {
  if (!condition) {
    throw new Error(`check-job-status FAILED: ${message}`);
  }
  console.log(`  ok: ${message}`);
}

function createTest() {
  const context = createGasContext();

  // Stateful stores replace the harness's write-only defaults.
  const scriptProperties = new Map();
  const scriptCache = new Map();

  const makeProperties = (store) => ({
    getProperty: (key) => (store.has(key) ? store.get(key) : null),
    setProperty: (key, value) => {
      store.set(key, String(value));
    },
    deleteProperty: (key) => {
      store.delete(key);
    },
    getKeys: () => Array.from(store.keys()),
  });

  const makeCache = (store) => ({
    get: (key) => (store.has(key) ? store.get(key) : null),
    put: (key, value) => {
      store.set(key, String(value));
    },
    remove: (key) => {
      store.delete(key);
    },
  });

  context.PropertiesService = {
    getScriptProperties: () => makeProperties(scriptProperties),
    getUserProperties: () => makeProperties(new Map()),
  };
  context.CacheService = { getScriptCache: () => makeCache(scriptCache) };

  let uuidCounter = 0;
  context.Utilities.getUuid = () => "job-" + (++uuidCounter);

  // Sidebar launcher recorder. createHtmlOutputFromFile succeeds for every
  // name so the primary "src/jobSidebar" candidate is exercised.
  const sidebarOpens = [];
  const appendedChunks = [];
  context.HtmlService = {
    createHtmlOutputFromFile: (name) => {
      const output = {
        __name: name,
        setTitle: () => output,
        append: (chunk) => {
          appendedChunks.push(chunk);
          return output;
        },
      };
      return output;
    },
  };
  const baseSpreadsheet = context.SpreadsheetApp.getActiveSpreadsheet;
  context.SpreadsheetApp = {
    getActiveSpreadsheet: baseSpreadsheet,
    getUi: () => ({
      showSidebar: (output) => {
        sidebarOpens.push(output.__name);
      },
      alert: () => "OK",
    }),
  };

  loadFiles(context, collectTypeScriptFiles(ignoredDirs));

  return { context, scriptProperties, scriptCache, sidebarOpens, appendedChunks };
}

function main() {
  console.log("job status store lifecycle:");
  const t = createTest();
  const context = t.context;

  assert(typeof context.startJobStatus === "function", "startJobStatus is global");
  assert(typeof context.getJobStatus === "function", "getJobStatus is global");
  assert(typeof context.updateJobStatusProgress === "function", "updateJobStatusProgress is global");
  assert(typeof context.completeJobStatus === "function", "completeJobStatus is global");
  assert(typeof context.failJobStatus === "function", "failJobStatus is global");
  assert(typeof context.cancelJobStatus === "function", "cancelJobStatus is global");
  assert(typeof context.clearJobStatus === "function", "clearJobStatus is global");
  assert(typeof context.openJobSidebar === "function", "openJobSidebar is global");

  // 1) Start: creates a running record readable through the store.
  const jobId = context.startJobStatus("generate", {
    title: "Generate CoMapeo Category",
    cancellable: false,
  });
  assert(typeof jobId === "string" && jobId.indexOf("job-") === 0, "startJobStatus returns a fresh job id");
  let state = context.getJobStatus(jobId);
  assert(state !== null, "getJobStatus returns the started job");
  assert(state.status === "running", "job starts as running");
  assert(state.percent === 0, "job starts at 0 percent");
  assert(state.operation === "generate", "operation recorded");
  assert(state.title === "Generate CoMapeo Category", "title recorded");
  assert(state.cancellable === false, "cancellable flag recorded");
  assert(Array.isArray(state.log) && state.log.length === 1, "start writes one log line");

  // 2) Progress: patch merges fields, clamps percent, appends to the log.
  context.updateJobStatusProgress(jobId, {
    percent: 40,
    phase: "Linting",
    detail: "checking sheets",
    logMessage: "lint ok",
  });
  state = context.getJobStatus(jobId);
  assert(state.percent === 40, "percent patch applied");
  assert(state.phase === "Linting", "phase patch applied");
  assert(state.detail === "checking sheets", "detail patch applied");
  assert(state.log.length === 2 && state.log[1].msg === "lint ok", "log line appended");

  context.updateJobStatusProgress(jobId, { percent: 150 });
  assert(context.getJobStatus(jobId).percent === 100, "percent clamped high");

  context.updateJobStatusProgress(jobId, { percent: -5 });
  assert(context.getJobStatus(jobId).percent === 0, "percent clamped low");

  // 3) Log tail cap: the store keeps only the newest entries.
  for (let i = 0; i < 40; i++) {
    context.updateJobStatusProgress(jobId, { logMessage: "entry " + i });
  }
  state = context.getJobStatus(jobId);
  assert(state.log.length === 30, "log tail capped at 30 entries");
  assert(state.log[state.log.length - 1].msg === "entry 39", "newest log entry kept");

  // 4) Terminal locking: patches after completion are ignored.
  context.completeJobStatus(jobId, "Configuration generated");
  state = context.getJobStatus(jobId);
  assert(state.status === "completed", "completeJobStatus marks completed");
  assert(state.percent === 100, "completion sets percent to 100");
  assert(state.summary === "Configuration generated", "summary stored");
  context.updateJobStatusProgress(jobId, { percent: 1, phase: "zombie" });
  state = context.getJobStatus(jobId);
  assert(state.percent === 100, "progress ignored after completion");
  assert(state.phase !== "zombie", "phase ignored after completion");

  // 5) Cache-eviction fallback: terminal state survives cache removal.
  t.scriptCache.clear();
  state = context.getJobStatus(jobId);
  assert(state !== null && state.status === "completed", "terminal state readable from properties after cache eviction");
  assert(state.summary === "Configuration generated", "summary survives cache eviction");
  const mirrorKey = "CMJ_JOB_" + jobId;
  assert(t.scriptProperties.has(mirrorKey), "terminal state mirrored to ScriptProperties");

  // 6) Failure path and cancellable flag visibility.
  const failJobId = context.startJobStatus("translation", { cancellable: true });
  context.failJobStatus(failJobId, "Translator threw");
  const failState = context.getJobStatus(failJobId);
  assert(failState.status === "error", "failJobStatus marks error");
  assert(failState.error === "Translator threw", "error message stored");
  assert(failState.cancellable === true, "cancellable jobs expose the flag to the sidebar");
  context.cancelJobStatus(failJobId, "Cancelled by user");
  assert(context.getJobStatus(failJobId).status === "error", "cancel after terminal state is ignored");

  // 7) Cancelled terminal state.
  const cancelJobId = context.startJobStatus("translation", { cancellable: true });
  context.cancelJobStatus(cancelJobId, "Translation cancelled. Partial translations were kept.");
  const cancelState = context.getJobStatus(cancelJobId);
  assert(cancelState.status === "cancelled", "cancelJobStatus marks cancelled");
  assert(
    cancelState.summary === "Translation cancelled. Partial translations were kept.",
    "cancel summary stored",
  );

  // 8) Clear lifecycle: cache + properties both emptied.
  context.clearJobStatus(cancelJobId);
  assert(context.getJobStatus(cancelJobId) === null, "cleared job unreadable");
  assert(!t.scriptCache.has("CMJ_JOB_" + cancelJobId), "cache entry removed");
  assert(!t.scriptProperties.has("CMJ_JOB_" + cancelJobId), "properties entry removed");

  // 9) Unknown / empty ids: reads and progress stay no-ops, but ending a
  // job whose state is unreadable synthesizes a minimal terminal record
  // (RISK 2) so the sidebar always reaches a terminal state.
  assert(context.getJobStatus("nope") === null, "unknown id returns null");
  assert(context.getJobStatus("") === null, "empty id returns null");
  context.updateJobStatusProgress("", { percent: 50 });
  context.updateJobStatusProgress("nope", { percent: 50 });
  context.clearJobStatus("");
  assert(context.getJobStatus("") === null, "empty id still returns null after no-op calls");
  context.completeJobStatus("nope");
  const synthState = context.getJobStatus("nope");
  assert(
    synthState !== null && synthState.status === "completed",
    "end with unreadable state synthesizes a terminal record",
  );
  assert(synthState.operation === "unknown", "synthesized record marks operation unknown");
  assert(synthState.percent === 100, "synthesized completion sits at 100 percent");
  assert(
    Array.isArray(synthState.log) && synthState.log.length === 1,
    "synthesized record keeps a minimal one-line log",
  );
  assert(t.scriptProperties.has("CMJ_JOB_nope"), "synthesized record mirrored to ScriptProperties");
  context.failJobStatus("");
  context.cancelJobStatus("nope");
  assert(
    context.getJobStatus("nope").status === "completed",
    "later ends ignore the synthesized terminal record",
  );

  // 10) Sidebar launcher: file output + injected job id + showSidebar.
  const sidebarJobId = context.startJobStatus("import", { title: "Import Category File" });
  context.openJobSidebar(sidebarJobId);
  assert(t.sidebarOpens.length === 1, "openJobSidebar calls showSidebar once");
  assert(t.sidebarOpens[0] === "src/jobSidebar", "sidebar loaded from src/jobSidebar.html");
  assert(
    t.appendedChunks.some((chunk) => String(chunk).indexOf('window.SIDEBAR_JOB_ID = "job-') !== -1),
    "bootstrap script injects the job id",
  );

  // 11) Poll payload must be JSON-safe for google.script.run.
  const payload = context.getJobStatus(sidebarJobId);
  const roundTrip = JSON.parse(JSON.stringify(payload));
  assert(roundTrip.jobId === sidebarJobId, "job state survives JSON round-trip");

  // 12) RISK 2 regression: CacheService evicts a RUNNING job mid-run.
  // Running states are cache-only (never mirrored), so end/cancel used to
  // read null and silently no-op, leaving the sidebar polling "no data".
  console.log("eviction-mid-run terminal synthesis:");
  const e = createTest();
  const ectx = e.context;

  const evictId = ectx.startJobStatus("translation", { cancellable: true });
  ectx.updateJobStatusProgress(evictId, { percent: 50, phase: "Translating" });
  e.scriptCache.clear(); // simulate cache eviction mid-run
  assert(ectx.getJobStatus(evictId) === null, "evicted running state is unreadable before end");
  ectx.completeJobStatus(evictId, "Finished after eviction");
  const evictState = ectx.getJobStatus(evictId);
  assert(evictState !== null, "completeJobStatus synthesizes a record after eviction");
  assert(evictState.status === "completed", "synthesized record carries the terminal status");
  assert(evictState.percent === 100, "synthesized completion sits at 100 percent");
  assert(evictState.summary === "Finished after eviction", "synthesized record keeps the summary");
  assert(
    e.scriptProperties.has("CMJ_JOB_" + evictId),
    "synthesized record mirrored to ScriptProperties",
  );

  const cancelEvictId = ectx.startJobStatus("translation", { cancellable: true });
  e.scriptCache.clear();
  ectx.cancelJobStatus(cancelEvictId, "Cancelled after eviction");
  assert(
    ectx.getJobStatus(cancelEvictId).status === "cancelled",
    "cancelJobStatus synthesizes a cancelled record after eviction",
  );

  const failEvictId = ectx.startJobStatus("translation", {});
  e.scriptCache.clear();
  ectx.failJobStatus(failEvictId, "Boom after eviction");
  assert(
    ectx.getJobStatus(failEvictId).status === "error",
    "failJobStatus synthesizes an error record after eviction",
  );

  // Progress updates on the evicted job stay silent no-ops.
  const progressEvictId = ectx.startJobStatus("translation", {});
  e.scriptCache.clear();
  ectx.updateJobStatusProgress(progressEvictId, { percent: 90, phase: "zombie" });
  assert(ectx.getJobStatus(progressEvictId) === null, "progress on evicted running job stays a no-op");

  // 13) RISK 1 regression: startJobStatus sweeps stale mirrored terminal
  // records, caps the total, leaves unparseable mirrors, and never touches
  // a concurrently running job (running states are cache-only).
  console.log("stale-record sweep on start:");
  const s = createTest();
  const sctx = s.context;

  const threeHoursAgo = Date.now() - 3 * 60 * 60 * 1000;
  const staleRecord = {
    jobId: "stale-job",
    operation: "generate",
    title: "Old run",
    status: "completed",
    percent: 100,
    phase: "completed",
    detail: "",
    log: [{ t: threeHoursAgo, msg: "old", level: "info" }],
    cancellable: false,
    startedAt: threeHoursAgo,
    touchedAt: threeHoursAgo,
  };
  s.scriptProperties.set("CMJ_JOB_stale-job", JSON.stringify(staleRecord));
  s.scriptProperties.set("CMJ_JOB_broken-mirror", "not-json{");
  for (let i = 0; i < 25; i++) {
    s.scriptProperties.set(
      "CMJ_JOB_fresh-" + i,
      JSON.stringify({
        ...staleRecord,
        jobId: "fresh-" + i,
        title: "Fresh " + i,
        touchedAt: Date.now() + i, // fresh-24 is the newest
      }),
    );
  }

  const runId = sctx.startJobStatus("generate", { title: "New run" });
  assert(!s.scriptProperties.has("CMJ_JOB_stale-job"), "sweep removes terminal mirrors older than the threshold");
  assert(!s.scriptProperties.has("CMJ_JOB_broken-mirror"), "sweep drops unparseable mirrors");
  const remaining = Array.from(s.scriptProperties.keys()).filter(
    (key) => key.indexOf("CMJ_JOB_") === 0,
  );
  assert(remaining.length === 20, "sweep caps mirrored terminal records at the cap");
  assert(s.scriptProperties.has("CMJ_JOB_fresh-24"), "sweep keeps the newest mirrors");
  assert(!s.scriptProperties.has("CMJ_JOB_fresh-0"), "sweep drops the oldest surplus mirrors");
  const runState = sctx.getJobStatus(runId);
  assert(runState !== null && runState.status === "running", "sweep preserves the concurrently running job");
  assert(
    remaining.indexOf("CMJ_JOB_" + runId) === -1,
    "running state stays cache-only, untouched by the sweep",
  );
  // A second start (another sweep pass) keeps the store at the cap.
  sctx.startJobStatus("generate", { title: "Another run" });
  const remainingAfterSecondSweep = Array.from(s.scriptProperties.keys()).filter(
    (key) => key.indexOf("CMJ_JOB_") === 0,
  );
  assert(
    remainingAfterSecondSweep.length <= 21,
    "repeated sweeps keep the mirrored store capped",
  );

  console.log("check-job-status PASSED");
}

try {
  main();
} catch (error) {
  console.error(error.message);
  process.exit(1);
}
