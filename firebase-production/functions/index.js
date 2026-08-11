"use strict";

const { timingSafeEqual } = require("node:crypto");
const { getApps, initializeApp } = require("firebase-admin/app");
const { FieldValue, getFirestore } = require("firebase-admin/firestore");
const { getMessaging } = require("firebase-admin/messaging");
const { logger, setGlobalOptions } = require("firebase-functions/v2");
const { onDocumentCreated } = require("firebase-functions/v2/firestore");
const { onRequest } = require("firebase-functions/v2/https");
const { onSchedule } = require("firebase-functions/v2/scheduler");
const { defineSecret } = require("firebase-functions/params");
const { GoogleAuth } = require("google-auth-library");
const {
  PRODUCTION_PROJECT_ID,
  createPublisher
} = require("./lib/publisher");
const { MAX_REQUEST_BYTES, createSubmissionService } = require("./lib/submission");
const { createNotificationService } = require("./lib/notification-service");
const {
  createHumetroBoardPostsProvider,
  createHumetroLatestPostProvider
} = require("./lib/humetro-client");
const { createSheetsGateway } = require("./lib/sheets-sync");
const { createSynchronizer } = require("./lib/synchronizer");

const REGION = "asia-northeast3";
const PUBLISHER_ADMIN_TOKEN = defineSecret("BWA_PUBLISHER_TOKEN");
const HUMETRO_BRIDGE_TOKEN = defineSecret("HUMETRO_BRIDGE_TOKEN");
const PUBLIC_CORS = ["https://somyun.github.io"];
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";

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

function getPublisher() {
  assertProductionProject();
  return createPublisher({
    firestore: getFirestore(getAdminApp()),
    serverTimestamp: () => FieldValue.serverTimestamp(),
    logger
  });
}

function getAdminApp() {
  return getApps().find((app) => app.name === "[DEFAULT]") || initializeApp();
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

let notificationService;
function getNotificationService() {
  if (!notificationService) {
    notificationService = createNotificationService({
      firestore: getFirestoreClient(),
      messaging: getMessaging(getAdminApp()),
      serverTimestamp: () => FieldValue.serverTimestamp(),
      latestPostProvider: createHumetroLatestPostProvider({
        bridgeTokenProvider: () => HUMETRO_BRIDGE_TOKEN.value()
      }),
      boardPostsProvider: createHumetroBoardPostsProvider({
        bridgeTokenProvider: () => HUMETRO_BRIDGE_TOKEN.value()
      }),
      logger
    });
  }
  return notificationService;
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

function hasValidAdminToken(request) {
  const privateHeader = String(request.get("x-bwa-publisher-token") || "");
  const authorization = String(request.get("authorization") || "");
  const prefix = "Bearer ";
  const candidate = (privateHeader || (authorization.startsWith(prefix) ? authorization.slice(prefix.length) : "")).trim();
  if (!candidate) return false;
  const supplied = Buffer.from(candidate, "utf8");
  const expected = Buffer.from(String(PUBLISHER_ADMIN_TOKEN.value()).trim(), "utf8");
  return supplied.length === expected.length && supplied.length > 0 && timingSafeEqual(supplied, expected);
}

function parseBody(request) {
  if (!request.body) return {};
  if (Buffer.isBuffer(request.body)) return JSON.parse(request.body.toString("utf8"));
  if (typeof request.body === "string") return JSON.parse(request.body);
  return request.body;
}

function assertPublicOrigin(request) {
  const origin = String(request.get("origin") || "");
  const allowed = origin === "https://somyun.github.io";
  if (!allowed) throw new Error("ORIGIN_NOT_ALLOWED");
}

function assertBodySize(request) {
  const contentLength = Number(request.get("content-length") || 0);
  if (Number.isFinite(contentLength) && contentLength > MAX_REQUEST_BYTES) {
    throw new Error("SUBMISSION_PAYLOAD_TOO_LARGE");
  }
  const bodyBytes = Buffer.byteLength(JSON.stringify(parseBody(request)), "utf8");
  if (bodyBytes > MAX_REQUEST_BYTES) throw new Error("SUBMISSION_PAYLOAD_TOO_LARGE");
}

function statusForError(error) {
  const message = String(error?.message || "");
  if (/NOT_FOUND/.test(message)) return 404;
  if (/UNAUTHORIZED_DEVICE/.test(message)) return 403;
  if (/CONFLICT|SYNC_IN_PROGRESS|STALE/.test(message)) return 409;
  if (/INACTIVE_DEVICE|FCM_TOKEN_UNREGISTERED/.test(message)) return 409;
  if (/RATE_LIMITED/.test(message)) return 429;
  if (/HUMETRO_/.test(message)) return 502;
  if (/DISABLED/.test(message)) return 503;
  if (/INVALID|UNSUPPORTED|MISMATCH|DUPLICATE|ORIGIN_NOT_ALLOWED|FORM_NOT_ALLOWED|PAYLOAD_TOO_LARGE/.test(message)) return 400;
  return 500;
}

function publicEndpoint(handler, release = "t5-async-save", { secrets = [] } = {}) {
  return onRequest({
    cors: PUBLIC_CORS,
    timeoutSeconds: 60,
    secrets,
    labels: { "bwa-release": release }
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
      const result = await handler(parseBody(request));
      response.status(200).json({ ok: true, result });
    } catch (error) {
      const errorMessage = String(error?.message || "UNKNOWN_ERROR").slice(0, 200);
      logger.warn("Public T5 endpoint rejected request", { error: errorMessage });
      response.status(statusForError(error)).json({ ok: false, error: errorMessage });
    }
  });
}

function manualAdmin(handler) {
  return onRequest({
    secrets: [PUBLISHER_ADMIN_TOKEN],
    cors: false,
    labels: { "bwa-release": "t5-async-save" }
  }, async (request, response) => {
    response.set("Cache-Control", "no-store");
    if (request.method !== "POST") {
      response.set("Allow", "POST").status(405).json({ error: "METHOD_NOT_ALLOWED" });
      return;
    }
    if (!hasValidAdminToken(request)) {
      response.status(401).json({ error: "UNAUTHORIZED" });
      return;
    }
    try {
      const result = await handler(parseBody(request));
      response.status(200).json({ ok: true, result });
    } catch (error) {
      logger.error("Manual T5 admin action failed", { message: error.message });
      response.status(statusForError(error)).json({ ok: false, error: error.message });
    }
  });
}

function manualPublisher(handler) {
  return manualAdmin((body) => handler(getPublisher(), body));
}

exports.publishFormList = manualPublisher((publisher, body) =>
  publisher.publishFormList({ force: Boolean(body.force) }));

exports.publishForm = manualPublisher((publisher, body) =>
  publisher.publishForm({ sheetName: body.sheetName, force: Boolean(body.force) }));

exports.publishAllChangedForms = manualPublisher((publisher, body) =>
  publisher.publishAllChangedForms({ force: Boolean(body.force) }));

exports.publishAllChangedFormsScheduled = onSchedule({
  schedule: "every 5 minutes",
  timeZone: "Asia/Seoul",
  retryCount: 0
}, async () => {
  await getPublisher().publishAllChangedForms();
});

exports.submitMeasurements = publicEndpoint((body) => getSubmissionService().submit(body));

exports.getMeasurementSubmission = publicEndpoint((body) =>
  getSubmissionService().getStatus(body.idempotencyKey));

exports.registerNotificationDevice = publicEndpoint((body) =>
  getNotificationService().registerDevice(body), "t7-notification-health");

exports.setNotificationDeviceActive = publicEndpoint((body) =>
  getNotificationService().setDeviceActive(body), "t7-notification-health");

exports.getNotificationDeviceStatus = publicEndpoint((body) =>
  getNotificationService().getStatus(body), "t7-notification-health");

exports.acknowledgeNotification = publicEndpoint((body) =>
  getNotificationService().acknowledge(body), "t7-notification-health");

exports.sendNotificationSelfTest = publicEndpoint((body) =>
  getNotificationService().sendSelfTest(body), "t11-notification-pages", {
    secrets: [HUMETRO_BRIDGE_TOKEN]
  });

exports.sendNotificationHeartbeatScheduled = onSchedule({
  schedule: "0 9 * * *",
  timeZone: "Asia/Seoul",
  retryCount: 0,
  labels: { "bwa-release": "t7-notification-health" }
}, async () => {
  const result = await getNotificationService().sendHeartbeatAll();
  logger.info("Daily notification heartbeat completed", result);
});

exports.sendHappyHugetherNotificationsScheduled = onSchedule({
  schedule: "every 10 minutes",
  timeZone: "Asia/Seoul",
  retryCount: 0,
  secrets: [HUMETRO_BRIDGE_TOKEN],
  labels: { "bwa-release": "t14-firebase-notifications" }
}, async () => {
  const gate = await getFirestoreClient().collection("systemConfig").doc("notificationDispatch").get();
  if (!gate.exists || gate.get("enabled") !== true) {
    logger.info("Happy Hugether notification dispatch is disabled");
    return { skipped: true, reason: "NOTIFICATION_DISPATCH_DISABLED" };
  }
  const result = await getNotificationService().sendLatestBoardPostIfNew();
  logger.info("Happy Hugether notification check completed", result);
});

exports.syncMeasurementSubmission = onDocumentCreated({
  document: "measurementSubmissions/{submissionId}",
  retry: true,
  maxInstances: 1,
  concurrency: 1,
  labels: { "bwa-release": "t5-async-save" }
}, async (event) => {
  const submissionId = event.params.submissionId;
  const synchronizer = await getSynchronizer();
  return synchronizer.syncSubmission(submissionId, event.id, { throwRetryable: true });
});

exports.retryMeasurementSubmission = manualAdmin(async (body) => {
  const synchronizer = await getSynchronizer();
  const submissionId = String(body.submissionId || "").trim();
  if (!submissionId) throw new Error("INVALID_SUBMISSION_ID");
  return synchronizer.syncSubmission(submissionId, `manual-${Date.now()}`);
});

exports.setSubmissionGate = manualAdmin((body) =>
  getSubmissionService().setGate(body.enabled === true, {
    allowedFormKeys: body.allowedFormKeys,
    allowedSheetNames: body.allowedSheetNames
  }));

exports.setNotificationDispatchGate = manualAdmin(async (body) => {
  const enabled = body.enabled === true;
  await getFirestoreClient().collection("systemConfig").doc("notificationDispatch").set({
    enabled,
    updatedAt: FieldValue.serverTimestamp()
  }, { merge: true });
  return { enabled };
});
