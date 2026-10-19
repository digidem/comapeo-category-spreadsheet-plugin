/// <reference path="./types.ts" />
/// <reference path="./loggingHelpers.ts" />

/**
 * Job status store for long-running operations (issue #57).
 *
 * Global variables do not persist between google.script.run calls, so a
 * running operation records its progress under a job id in CacheService
 * (hot path, 6h TTL) with terminal states mirrored to ScriptProperties so
 * the final result survives cache eviction. The sidebar opened by
 * openJobSidebar() polls getJobStatus(jobId) — those polls run as separate
 * executions and only ever read the store, never module state.
 */

type JobStatusValue = "running" | "completed" | "error" | "cancelled";

interface JobLogEntry {
  /** Epoch milliseconds when the line was recorded. */
  t: number;
  msg: string;
  level?: "info" | "warn" | "error";
}

interface JobStatusState {
  jobId: string;
  operation: string;
  title: string;
  status: JobStatusValue;
  /** 0..100, clamped. */
  percent: number;
  phase: string;
  detail: string;
  log: JobLogEntry[];
  cancellable: boolean;
  startedAt: number;
  touchedAt: number;
  summary?: string;
  error?: string;
}

interface JobStatusOptions {
  title?: string;
  detail?: string;
  /** Sidebar Cancel is hardwired to requestTranslationCancel — translation jobs only. */
  cancellable?: boolean;
}

interface JobProgressPatch {
  percent?: number;
  phase?: string;
  detail?: string;
  logMessage?: string;
  logLevel?: "info" | "warn" | "error";
}

/**
 * CacheService maximum TTL. Longer than any supported operation, so the
 * state never expires mid-run.
 */
const JOB_CACHE_TTL_SECONDS = 21600;
const JOB_KEY_PREFIX = "CMJ_JOB_";
// Keeps the JSON blob far below the 100KB cache value limit even with
// chatty phases; the sidebar only renders a tail anyway.
const JOB_LOG_MAX_ENTRIES = 30;
// ScriptProperties is a ~500KB store shared with driveService/utils/
// performanceMetrics. Without a sweep every terminal job would permanently
// leave its mirrored record (log included) there, so startJobStatus drops
// mirrors older than an hour and caps the total count. Running states are
// cache-only and never appear in ScriptProperties, so a sweep can never
// break a concurrently running job's read fallback.
const JOB_SWEEP_MAX_AGE_MS = 60 * 60 * 1000;
const JOB_SWEEP_MAX_RECORDS = 20;

function getJobStatusLogger() {
  return typeof getScopedLogger === "function"
    ? getScopedLogger("JobStatus")
    : console;
}

function buildJobKey(jobId: string): string {
  return JOB_KEY_PREFIX + String(jobId);
}

/**
 * Unique job id. Utilities.getUuid is the GAS source; the fallback keeps
 * job creation working in environments whose Utilities mock lacks it (the
 * id only has to be unique within this spreadsheet's stores).
 */
function createJobId(): string {
  if (typeof Utilities !== "undefined" && typeof Utilities.getUuid === "function") {
    return Utilities.getUuid();
  }
  return `job-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function writeJobState(state: JobStatusState): void {
  const key = buildJobKey(state.jobId);
  const serialized = JSON.stringify(state);
  try {
    CacheService.getScriptCache().put(key, serialized, JOB_CACHE_TTL_SECONDS);
  } catch (error) {
    getJobStatusLogger().warn("JobStatus cache write failed:", error);
  }
  // Terminal states are mirrored to ScriptProperties so the final result
  // survives cache eviction and is still readable when the sidebar polls.
  if (state.status !== "running") {
    try {
      PropertiesService.getScriptProperties().setProperty(key, serialized);
    } catch (error) {
      getJobStatusLogger().warn("JobStatus properties write failed:", error);
    }
  }
}

function readJobState(jobId: string): JobStatusState | null {
  const key = buildJobKey(jobId);
  let raw: string | null = null;
  try {
    raw = CacheService.getScriptCache().get(key);
  } catch (error) {
    getJobStatusLogger().warn("JobStatus cache read failed:", error);
  }
  if (raw === null || raw === undefined) {
    try {
      raw = PropertiesService.getScriptProperties().getProperty(key);
    } catch (error) {
      getJobStatusLogger().warn("JobStatus properties read failed:", error);
    }
  }
  if (!raw) {
    return null;
  }
  try {
    const parsed = JSON.parse(raw) as JobStatusState;
    return parsed && parsed.jobId ? parsed : null;
  } catch (error) {
    getJobStatusLogger().warn("JobStatus state parse failed:", error);
    return null;
  }
}

function appendJobLogLine(state: JobStatusState, entry: JobLogEntry): void {
  state.log.push(entry);
  if (state.log.length > JOB_LOG_MAX_ENTRIES) {
    state.log.splice(0, state.log.length - JOB_LOG_MAX_ENTRIES);
  }
}

/**
 * Removes stale mirrored terminal records from ScriptProperties so the
 * shared property store does not fill up with completed jobs nobody will
 * read again (clearJobStatus has no production caller). Only CMJ_JOB_ keys
 * are touched; a concurrently running job keeps its state in the cache, so
 * this can never break its read fallback. Failures are logged and ignored —
 * housekeeping must never prevent a job from starting.
 */
function sweepStaleJobRecords(): void {
  try {
    const props = PropertiesService.getScriptProperties();
    if (typeof props.getKeys !== "function") {
      return;
    }
    const keys = props
      .getKeys()
      .filter((key) => key.indexOf(JOB_KEY_PREFIX) === 0);
    if (keys.length === 0) {
      return;
    }
    const now = Date.now();
    const records: { key: string; touchedAt: number }[] = [];
    for (const key of keys) {
      let touchedAt = NaN;
      try {
        const raw = props.getProperty(key);
        const parsed = raw ? (JSON.parse(raw) as JobStatusState) : null;
        if (parsed && typeof parsed.touchedAt === "number") {
          touchedAt = parsed.touchedAt;
        }
      } catch (parseError) {
        getJobStatusLogger().warn(`JobStatus mirror ${key} unparseable:`, parseError);
      }
      if (!isFinite(touchedAt)) {
        // readJobState could never use this mirror again — drop it.
        props.deleteProperty(key);
        continue;
      }
      records.push({ key, touchedAt });
    }
    records.sort((a, b) => b.touchedAt - a.touchedAt);
    for (let i = 0; i < records.length; i++) {
      const tooOld = now - records[i].touchedAt > JOB_SWEEP_MAX_AGE_MS;
      const overCap = i >= JOB_SWEEP_MAX_RECORDS;
      if (tooOld || overCap) {
        props.deleteProperty(records[i].key);
      }
    }
  } catch (error) {
    getJobStatusLogger().warn("JobStatus sweep skipped:", error);
  }
}

/**
 * Creates a running job record and returns its id. The id must be passed to
 * updateJobStatusProgress / completeJobStatus / failJobStatus by the running
 * operation and to openJobSidebar + getJobStatus by the sidebar.
 */
function startJobStatus(operation: string, options?: JobStatusOptions): string {
  sweepStaleJobRecords();
  const jobId = createJobId();
  const now = Date.now();
  const state: JobStatusState = {
    jobId,
    operation: String(operation || "operation"),
    title: String((options && options.title) || operation || "Working"),
    status: "running",
    percent: 0,
    phase: "Starting",
    detail: String((options && options.detail) || ""),
    log: [],
    cancellable: Boolean(options && options.cancellable),
    startedAt: now,
    touchedAt: now,
  };
  appendJobLogLine(state, { t: now, msg: `${state.title}: started`, level: "info" });
  if (state.detail) {
    appendJobLogLine(state, { t: now, msg: state.detail, level: "info" });
  }
  writeJobState(state);
  getJobStatusLogger().info(`[JOB] ${state.operation} started (${jobId})`);
  return jobId;
}

/**
 * Merges a progress patch into the running job's state. Silent no-op when
 * the job id is unknown (store evicted) or the job already ended — the
 * operation must never crash because progress reporting failed.
 */
function updateJobStatusProgress(jobId: string, patch?: JobProgressPatch): void {
  if (!jobId) {
    return;
  }
  try {
    const state = readJobState(jobId);
    if (!state || state.status !== "running") {
      return;
    }
    const now = Date.now();
    if (patch && typeof patch.percent === "number" && isFinite(patch.percent)) {
      state.percent = Math.max(0, Math.min(100, Math.round(patch.percent)));
    }
    if (patch && typeof patch.phase === "string" && patch.phase) {
      state.phase = patch.phase;
    }
    if (patch && typeof patch.detail === "string") {
      state.detail = patch.detail;
    }
    if (patch && patch.logMessage) {
      appendJobLogLine(state, {
        t: now,
        msg: String(patch.logMessage),
        level: (patch.logLevel as JobLogEntry["level"]) || "info",
      });
    }
    state.touchedAt = now;
    writeJobState(state);
  } catch (error) {
    getJobStatusLogger().warn("JobStatus progress update failed:", error);
  }
}

/**
 * Creates a minimal terminal record for a job whose running state is no
 * longer readable — CacheService evicted a RUNNING job mid-run (running
 * states are never mirrored, so the ScriptProperties fallback is empty
 * too). Without this the sidebar would poll "no data" forever while the
 * toast already claimed completed/cancelled.
 */
function synthesizeTerminalJobState(
  jobId: string,
  status: JobStatusValue,
  message: string,
  fields?: { summary?: string; error?: string },
): void {
  const now = Date.now();
  const state: JobStatusState = {
    jobId,
    operation: "unknown",
    title: "Operation",
    status,
    percent: status === "completed" ? 100 : 0,
    phase: status,
    detail: "",
    log: [
      {
        t: now,
        msg: message,
        level:
          status === "error" ? "error" : status === "cancelled" ? "warn" : "info",
      },
    ],
    cancellable: false,
    startedAt: now,
    touchedAt: now,
  };
  if (fields && fields.summary) {
    state.summary = fields.summary;
  }
  if (fields && fields.error) {
    state.error = fields.error;
  }
  writeJobState(state);
  getJobStatusLogger().warn(
    `[JOB] ${jobId} ${status} (state unreadable; synthesized terminal record)`,
  );
}

function endJobStatus(
  jobId: string,
  status: JobStatusValue,
  message: string,
  fields?: { summary?: string; error?: string },
): void {
  if (!jobId) {
    return;
  }
  try {
    const state = readJobState(jobId);
    if (!state) {
      synthesizeTerminalJobState(jobId, status, message, fields);
      return;
    }
    if (state.status !== "running") {
      return;
    }
    const now = Date.now();
    state.status = status;
    if (status === "completed") {
      state.percent = 100;
    }
    if (fields && fields.summary) {
      state.summary = fields.summary;
    }
    if (fields && fields.error) {
      state.error = fields.error;
    }
    appendJobLogLine(state, {
      t: now,
      msg: message,
      level: status === "error" ? "error" : status === "cancelled" ? "warn" : "info",
    });
    state.touchedAt = now;
    writeJobState(state);
    getJobStatusLogger().info(`[JOB] ${state.operation} ${status} (${jobId})`);
  } catch (error) {
    getJobStatusLogger().warn("JobStatus end update failed:", error);
  }
}

function completeJobStatus(jobId: string, summary?: string): void {
  endJobStatus(jobId, "completed", summary || "Completed", { summary: summary || "Completed" });
}

function failJobStatus(jobId: string, errorMessage: string): void {
  endJobStatus(jobId, "error", errorMessage || "Failed", {
    error: errorMessage || "Failed",
  });
}

function cancelJobStatus(jobId: string, message?: string): void {
  endJobStatus(jobId, "cancelled", message || "Cancelled by user", {
    summary: message || "Cancelled by user",
  });
}

/**
 * Client-facing polling endpoint (google.script.run). Returns the job state
 * as a plain JSON-safe object, or null when the id is unknown.
 */
function getJobStatus(jobId: string): JobStatusState | null {
  if (!jobId) {
    return null;
  }
  return readJobState(String(jobId));
}

function clearJobStatus(jobId: string): void {
  if (!jobId) {
    return;
  }
  const key = buildJobKey(jobId);
  try {
    CacheService.getScriptCache().remove(key);
  } catch (error) {
    getJobStatusLogger().warn("JobStatus cache clear failed:", error);
  }
  try {
    PropertiesService.getScriptProperties().deleteProperty(key);
  } catch (error) {
    getJobStatusLogger().warn("JobStatus properties clear failed:", error);
  }
}

/**
 * Opens the progress sidebar for a job. The sidebar polls getJobStatus via
 * google.script.run until the job reaches a terminal state. Safe to call in
 * harness contexts without HtmlService — it logs and skips instead.
 */
function openJobSidebar(jobId: string): void {
  if (!jobId) {
    return;
  }
  try {
    if (typeof HtmlService === "undefined") {
      getJobStatusLogger().info("HtmlService unavailable; sidebar skipped");
      return;
    }
    const candidates = ["src/jobSidebar", "jobSidebar"];
    let output: GoogleAppsScript.HTML.HtmlOutput | null = null;
    let lastError: Error | null = null;
    for (const name of candidates) {
      try {
        output = HtmlService.createHtmlOutputFromFile(name);
        break;
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));
      }
    }
    if (!output) {
      throw lastError || new Error("jobSidebar.html not found");
    }
    output.setTitle("CoMapeo");
    output.append(
      "<script>window.SIDEBAR_JOB_ID = " +
        // JSON.stringify leaves "</script>" intact, which would terminate
        // the inline bootstrap early — escape every "<".
        JSON.stringify(String(jobId)).replace(/</g, "\\u003c") +
        ";</script>",
    );
    const ui = SpreadsheetApp.getUi();
    if (typeof ui.showSidebar !== "function") {
      getJobStatusLogger().info("showSidebar unavailable; sidebar skipped");
      return;
    }
    ui.showSidebar(output);
    getJobStatusLogger().info(`[JOB] sidebar opened for job ${jobId}`);
  } catch (error) {
    getJobStatusLogger().warn("Unable to open job sidebar:", error);
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (spreadsheet && typeof spreadsheet.toast === "function") {
        spreadsheet.toast(
          "Progress sidebar could not be opened; the operation continues.",
          "CoMapeo",
          5,
        );
      }
    } catch (toastError) {
      getJobStatusLogger().warn("Sidebar failure toast skipped:", toastError);
    }
  }
}
