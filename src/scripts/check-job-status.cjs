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

  // 9) Unknown / empty ids are safe no-ops.
  assert(context.getJobStatus("nope") === null, "unknown id returns null");
  assert(context.getJobStatus("") === null, "empty id returns null");
  context.updateJobStatusProgress("", { percent: 50 });
  context.updateJobStatusProgress("nope", { percent: 50 });
  context.completeJobStatus("nope");
  context.failJobStatus("");
  context.cancelJobStatus("nope");
  context.clearJobStatus("");
  assert(context.getJobStatus("nope") === null, "unknown id still returns null after no-op calls");

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

  console.log("check-job-status PASSED");
}

try {
  main();
} catch (error) {
  console.error(error.message);
  process.exit(1);
}
