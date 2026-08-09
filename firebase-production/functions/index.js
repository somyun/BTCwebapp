"use strict";

const { getApps, initializeApp } = require("firebase-admin/app");
const { FieldValue, getFirestore } = require("firebase-admin/firestore");
const { logger, setGlobalOptions } = require("firebase-functions/v2");
const { onSchedule } = require("firebase-functions/v2/scheduler");
const { PRODUCTION_PROJECT_ID, createPublisher } = require("./lib/publisher");

const REGION = "asia-northeast3";

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

exports.publishAllChangedFormsScheduled = onSchedule({
  schedule: "every 5 minutes",
  timeZone: "Asia/Seoul",
  retryCount: 0,
  labels: { "bwa-release": "production-read-cache" }
}, async () => {
  await getPublisher().publishAllChangedForms();
});
