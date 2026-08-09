"use strict";

const { timingSafeEqual } = require("node:crypto");

const { getApps, initializeApp } = require("firebase-admin/app");
const { FieldValue, getFirestore } = require("firebase-admin/firestore");
const { logger, setGlobalOptions } = require("firebase-functions/v2");
const { onDocumentCreated } = require("firebase-functions/v2/firestore");
const { onRequest } = require("firebase-functions/v2/https");
const { onSchedule } = require("firebase-functions/v2/scheduler");
const { defineSecret } = require("firebase-functions/params");
const { GoogleAuth } = require("google-auth-library");
const { PRODUCTION_PROJECT_ID, createPublisher } = require("./lib/publisher");
const {
  MAX_REQUEST_BYTES,
  createSheetsGateway,
  createSubmissionService,
  createSynchronizer
} = require("./lib/submission-sync");

const REGION = "asia-northeast3";
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const PUBLIC_CORS = ["https://somyun.github.io", /^http:\/\/(?:127\.0\.0\.1|localhost):\d+$/];
const SUBMISSION_ADMIN_TOKEN = defineSecret("BWA_PRODUCTION_SUBMISSION_ADMIN_TOKEN");

setGlobalOptions({
  region: REGION,
  memory: "256MiB",
  timeoutSeconds: 300,
  maxInstances: 1,
  concurrency: 1
});

function assertProductionProject() {
  const projectId = process.env.GCLOUD_PROJECT || process.env.GCP_PROJECT;
  if (projectId !== PRODUCTION_PROJECT_ID) {
    throw new Error(`PRODUCTION_PROJECT_GUARD_FAILED:${projectId || "missing"}`);
  }
}

function getAdminApp() {
  return getApps().find((app) => app.name === "[DEFAULT]") || initializeApp();
}

function getPublisher() {
  assertProductionProject();
  return createPublisher({
    firestore: getFirestore(getAdminApp()),
    serverTimestamp: () => FieldValue.serverTimestamp(),
    logger
  });
}

function getFirestoreClient() {
  assertProductionProject();
  return getFirestore(getAdminApp());
}

let submissionService;
function getSubmissionService() {
  if (!submissionService) {
    submissionService = createSubmissionService({
      firestore: getFirestoreClient(),
      serverTimestamp: () => FieldValue.serverTimestamp()
    });
  }
  return submissionService;
}

let synchronizerPromise;
async function getSynchronizer() {
  if (!synchronizerPromise) {
    synchronizerPromise = (async () => {
      const authClient = await new GoogleAuth({ scopes: [SHEETS_SCOPE] }).getClient();
      return createSynchronizer({
        firestore: getFirestoreClient(),
        serverTimestamp: () => FieldValue.serverTimestamp(),
        sheetsGateway: createSheetsGateway({ authClient }),
        publisher: getPublisher(),
        logger
      });
    })();
  }
  return synchronizerPromise;
}

function parseBody(request) {
  if (!request.body) return {};
  if (Buffer.isBuffer(request.body)) return JSON.parse(request.body.toString("utf8"));
  if (typeof request.body === "string") return JSON.parse(request.body);
  return request.body;
}

function assertPublicOrigin(request) {
  const origin = String(request.get("origin") || "");
  if (origin !== "https://somyun.github.io" &&
      !/^http:\/\/(?:127\.0\.0\.1|localhost):\d+$/.test(origin)) {
    throw new Error("ORIGIN_NOT_ALLOWED");
  }
}

function assertBodySize(request) {
  const contentLength = Number(request.get("content-length") || 0);
  if (Number.isFinite(contentLength) && contentLength > MAX_REQUEST_BYTES) {
    throw new Error("SUBMISSION_PAYLOAD_TOO_LARGE");
  }
  if (Buffer.byteLength(JSON.stringify(parseBody(request)), "utf8") > MAX_REQUEST_BYTES) {
    throw new Error("SUBMISSION_PAYLOAD_TOO_LARGE");
  }
}

function statusForError(error) {
  const message = String(error?.message || "");
  if (/NOT_FOUND/.test(message)) return 404;
  if (/CONFLICT|SYNC_IN_PROGRESS|STALE/.test(message)) return 409;
  if (/RATE_LIMITED/.test(message)) return 429;
  if (/DISABLED/.test(message)) return 503;
  if (/INVALID|UNSUPPORTED|MISMATCH|DUPLICATE|ORIGIN_NOT_ALLOWED|PAYLOAD_TOO_LARGE/.test(message)) return 400;
  return 500;
}

function publicEndpoint(handler) {
  return onRequest({
    cors: PUBLIC_CORS,
    timeoutSeconds: 60,
    labels: { "bwa-release": "production-async-save" }
  }, async (request, response) => {
    response.set("Cache-Control", "no-store");
    if (request.method !== "POST") {
      response.set("Allow", "POST").status(405).json({ ok: false, error: "METHOD_NOT_ALLOWED" });
      return;
    }
    try {
      assertProductionProject();
      assertPublicOrigin(request);
      assertBodySize(request);
      response.status(200).json({ ok: true, result: await handler(parseBody(request)) });
    } catch (error) {
      const errorMessage = String(error?.message || "UNKNOWN_ERROR").slice(0, 200);
      logger.warn("Production submission endpoint rejected request", { error: errorMessage });
      response.status(statusForError(error)).json({ ok: false, error: errorMessage });
    }
  });
}

function hasValidAdminToken(request) {
  const privateHeader = String(request.get("x-bwa-submission-token") || "");
  const authorization = String(request.get("authorization") || "");
  const candidate = (privateHeader || (authorization.startsWith("Bearer ")
    ? authorization.slice("Bearer ".length)
    : "")).trim();
  const expectedText = String(SUBMISSION_ADMIN_TOKEN.value() || "").trim();
  if (!candidate || !expectedText) return false;
  const supplied = Buffer.from(candidate, "utf8");
  const expected = Buffer.from(expectedText, "utf8");
  return supplied.length === expected.length && timingSafeEqual(supplied, expected);
}

function manualAdmin(handler) {
  return onRequest({
    secrets: [SUBMISSION_ADMIN_TOKEN],
    cors: false,
    labels: { "bwa-release": "production-async-save" }
  }, async (request, response) => {
    response.set("Cache-Control", "no-store");
    if (request.method !== "POST") {
      response.set("Allow", "POST").status(405).json({ ok: false, error: "METHOD_NOT_ALLOWED" });
      return;
    }
    if (!hasValidAdminToken(request)) {
      response.status(401).json({ ok: false, error: "UNAUTHORIZED" });
      return;
    }
    try {
      response.status(200).json({ ok: true, result: await handler(parseBody(request)) });
    } catch (error) {
      response.status(statusForError(error)).json({ ok: false, error: error.message });
    }
  });
}

exports.publishAllChangedFormsScheduled = onSchedule({
  schedule: "every 5 minutes",
  timeZone: "Asia/Seoul",
  retryCount: 0,
  labels: { "bwa-release": "production-read-cache" }
}, async () => {
  await getPublisher().publishAllChangedForms();
});

exports.submitMeasurements = publicEndpoint((body) => getSubmissionService().submit(body));

exports.getMeasurementSubmission = publicEndpoint((body) =>
  getSubmissionService().getStatus(body.idempotencyKey));

exports.syncMeasurementSubmission = onDocumentCreated({
  document: "measurementSubmissions/{submissionId}",
  retry: true,
  maxInstances: 1,
  concurrency: 1,
  labels: { "bwa-release": "production-async-save" }
}, async (event) => {
  const synchronizer = await getSynchronizer();
  return synchronizer.syncSubmission(event.params.submissionId, event.id, { throwRetryable: true });
});

exports.retryMeasurementSubmission = manualAdmin(async (body) => {
  const submissionId = String(body.submissionId || "").trim();
  if (!submissionId) throw new Error("INVALID_SUBMISSION_ID");
  const synchronizer = await getSynchronizer();
  return synchronizer.syncSubmission(submissionId, `manual-${Date.now()}`);
});

exports.setSubmissionGate = manualAdmin((body) =>
  getSubmissionService().setGate(body.enabled === true));
