"use strict";

const { createHash } = require("node:crypto");
const { hash } = require("./notification-service");

const PRODUCTION_SPREADSHEET_ID = "19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s";
const LEGACY_TOKEN_RANGE = "'FCM_Tokens'!A2:E";

function activeValue(value) {
  if (value === true) return true;
  return String(value || "").trim().toLowerCase() === "true";
}

function legacyDeviceId(token) {
  return `d_${createHash("sha256").update(String(token), "utf8").digest("base64url")}`;
}

function normalizeLegacyTokenRows(rows) {
  if (!Array.isArray(rows)) throw new Error("INVALID_LEGACY_TOKEN_ROWS");
  const byHash = new Map();
  let invalid = 0;
  let duplicate = 0;
  let inactive = 0;
  for (const row of rows) {
    const token = String(row?.[0] || "").trim();
    if (token.length < 20 || token.length > 4096) {
      invalid += 1;
      continue;
    }
    if (!activeValue(row?.[4])) {
      inactive += 1;
      continue;
    }
    const tokenHash = hash(token);
    if (byHash.has(tokenHash)) duplicate += 1;
    byHash.set(tokenHash, {
      token,
      tokenHash,
      deviceId: legacyDeviceId(token),
      userAgent: String(row?.[1] || "").slice(0, 512),
      lastUpdated: String(row?.[2] || "").slice(0, 100),
      keywords: String(row?.[3] || "").trim().slice(0, 500),
      active: true
    });
  }
  return {
    devices: [...byHash.values()],
    summary: { sourceRows: rows.length, activeUnique: byHash.size, inactive, invalid, duplicate }
  };
}

function createLegacyNotificationMigrator({ firestore, sheetsGateway, serverTimestamp }) {
  if (!firestore || !sheetsGateway || typeof sheetsGateway.readRange !== "function" ||
      typeof serverTimestamp !== "function") {
    throw new Error("INVALID_LEGACY_MIGRATION_DEPENDENCIES");
  }

  async function migrate({ dryRun = true } = {}) {
    const rows = await sheetsGateway.readRange(LEGACY_TOKEN_RANGE);
    const normalized = normalizeLegacyTokenRows(rows);
    const result = { ...normalized.summary, dryRun: dryRun !== false, imported: 0, updated: 0, protected: 0 };
    if (dryRun !== false) return result;

    for (const device of normalized.devices) {
      const deviceReference = firestore.collection("notificationDevices").doc(device.deviceId);
      const ownerReference = firestore.collection("notificationTokenOwners").doc(device.tokenHash);
      await firestore.runTransaction(async (transaction) => {
        const deviceSnapshot = await transaction.get(deviceReference);
        const ownerSnapshot = await transaction.get(ownerReference);
        if (ownerSnapshot.exists && ownerSnapshot.get("deviceId") !== device.deviceId) {
          result.protected += 1;
          return;
        }
        const timestamp = serverTimestamp();
        const document = {
          schemaVersion: 1,
          deviceId: device.deviceId,
          secretHash: null,
          token: device.token,
          tokenHash: device.tokenHash,
          keywords: device.keywords,
          userAgent: device.userAgent,
          active: true,
          legacy: true,
          migrationSource: "FCM_Tokens",
          legacyLastUpdated: device.lastUpdated,
          updatedAt: timestamp,
          tokenUpdatedAt: timestamp,
          lastFailureCode: null,
          lastFailureAt: null
        };
        if (deviceSnapshot.exists) {
          if (deviceSnapshot.get("legacy") !== true) {
            result.protected += 1;
            return;
          }
          transaction.set(deviceReference, document, { merge: true });
          result.updated += 1;
        } else {
          transaction.create(deviceReference, { ...document, registeredAt: timestamp, tokenRotationCount: 0 });
          result.imported += 1;
        }
        transaction.set(ownerReference, {
          tokenHash: device.tokenHash,
          deviceId: device.deviceId,
          legacy: true,
          updatedAt: timestamp
        });
      });
    }
    await firestore.collection("systemConfig").doc("legacyNotificationMigration").set({
      ...result,
      sourceSpreadsheetId: PRODUCTION_SPREADSHEET_ID,
      completedAt: serverTimestamp()
    }, { merge: true });
    return result;
  }

  return { migrate };
}

module.exports = {
  LEGACY_TOKEN_RANGE,
  PRODUCTION_SPREADSHEET_ID,
  activeValue,
  createLegacyNotificationMigrator,
  legacyDeviceId,
  normalizeLegacyTokenRows
};
