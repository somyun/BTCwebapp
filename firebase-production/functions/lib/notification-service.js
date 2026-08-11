"use strict";

const { createHash, randomUUID, timingSafeEqual } = require("node:crypto");

const DEVICE_ID_PATTERN = /^d_[A-Za-z0-9_-]{32,80}$/;
const DEVICE_SECRET_PATTERN = /^[A-Za-z0-9_-]{32,160}$/;
const EVENT_ID_PATTERN = /^[A-Za-z0-9._:-]{8,160}$/;
const RECEIPT_PHASES = new Set(["received", "shown", "clicked"]);
const UNREGISTERED_CODES = new Set([
  "messaging/registration-token-not-registered",
  "messaging/invalid-registration-token"
]);

function boundedString(value, maxLength, errorCode, { required = false } = {}) {
  const normalized = String(value ?? "").trim();
  if ((required && !normalized) || normalized.length > maxLength) throw new Error(errorCode);
  return normalized;
}

function hash(value) {
  return createHash("sha256").update(String(value), "utf8").digest("hex");
}

function verifyHash(value, expectedHex) {
  const supplied = Buffer.from(hash(value), "hex");
  const expected = Buffer.from(String(expectedHex || ""), "hex");
  return supplied.length === expected.length && expected.length > 0 && timingSafeEqual(supplied, expected);
}

function normalizeCredentials(payload = {}) {
  const deviceId = boundedString(payload.deviceId, 82, "INVALID_DEVICE_ID", { required: true });
  const deviceSecret = boundedString(payload.deviceSecret, 160, "INVALID_DEVICE_SECRET", { required: true });
  if (!DEVICE_ID_PATTERN.test(deviceId)) throw new Error("INVALID_DEVICE_ID");
  if (!DEVICE_SECRET_PATTERN.test(deviceSecret)) throw new Error("INVALID_DEVICE_SECRET");
  return { deviceId, deviceSecret };
}

function normalizeRegistration(payload = {}) {
  const credentials = normalizeCredentials(payload);
  const token = boundedString(payload.token, 4096, "INVALID_FCM_TOKEN", { required: true });
  if (token.length < 20) throw new Error("INVALID_FCM_TOKEN");
  return {
    ...credentials,
    token,
    tokenHash: hash(token),
    keywords: boundedString(payload.keywords, 500, "INVALID_KEYWORDS"),
    userAgent: boundedString(payload.userAgent, 512, "INVALID_USER_AGENT"),
    active: payload.active !== false
  };
}

function normalizeReceipt(payload = {}) {
  const credentials = normalizeCredentials(payload);
  const eventId = boundedString(payload.eventId, 160, "INVALID_EVENT_ID", { required: true });
  const phase = boundedString(payload.phase, 20, "INVALID_RECEIPT_PHASE", { required: true });
  const type = boundedString(payload.type || "notification", 40, "INVALID_NOTIFICATION_TYPE");
  if (!EVENT_ID_PATTERN.test(eventId)) throw new Error("INVALID_EVENT_ID");
  if (!RECEIPT_PHASES.has(phase)) throw new Error("INVALID_RECEIPT_PHASE");
  return { ...credentials, eventId, phase, type };
}

function matchesKeywords(postTitle, keywordString) {
  const title = String(postTitle || "").toLocaleLowerCase("ko-KR");
  const keywords = String(keywordString || "").trim();
  if (!keywords) return true;
  const orGroups = keywords.split(",").map((group) => group.trim()).filter(Boolean);
  return orGroups.some((group) =>
    group.split(/\s+/).filter(Boolean).every((term) => title.includes(term.toLocaleLowerCase("ko-KR"))));
}

function isUnregisteredError(error) {
  const code = String(error?.code || "");
  const message = String(error?.message || "");
  return UNREGISTERED_CODES.has(code) || /UNREGISTERED|registration-token-not-registered|not registered/i.test(message);
}

function publicStatus(data = {}) {
  return {
    registered: true,
    active: data.active === true,
    tokenFingerprint: data.tokenHash ? String(data.tokenHash).slice(-12) : null,
    keywords: String(data.keywords || ""),
    registeredAt: data.registeredAt || null,
    updatedAt: data.updatedAt || null,
    tokenUpdatedAt: data.tokenUpdatedAt || null,
    lastNotificationAcceptedAt: data.lastNotificationAcceptedAt || null,
    lastNotificationReceivedAt: data.lastNotificationReceivedAt || null,
    lastShownAt: data.lastShownAt || null,
    lastClickedAt: data.lastClickedAt || null,
    lastHeartbeatAcceptedAt: data.lastHeartbeatAcceptedAt || null,
    lastHeartbeatReceivedAt: data.lastHeartbeatReceivedAt || null,
    lastFailureCode: data.lastFailureCode || null,
    lastFailureAt: data.lastFailureAt || null
  };
}

function createNotificationService({
  firestore,
  messaging,
  serverTimestamp,
  latestPostProvider,
  boardPostsProvider = async () => [await latestPostProvider()],
  now = () => new Date(),
  createEventId = () => randomUUID(),
  logger = console
}) {
  if (!firestore || !messaging || !serverTimestamp || typeof latestPostProvider !== "function" ||
      typeof boardPostsProvider !== "function") {
    throw new Error("INVALID_NOTIFICATION_DEPENDENCIES");
  }

  const devices = firestore.collection("notificationDevices");
  const tokenOwners = firestore.collection("notificationTokenOwners");
  const notificationState = firestore.collection("systemConfig").doc("happyHugetherNotifications");

  async function authenticatedDevice(payload) {
    const credentials = normalizeCredentials(payload);
    const reference = devices.doc(credentials.deviceId);
    const snapshot = await reference.get();
    if (!snapshot.exists || !verifyHash(credentials.deviceSecret, snapshot.data().secretHash)) {
      throw new Error("UNAUTHORIZED_DEVICE");
    }
    return { ...credentials, reference, data: snapshot.data() };
  }

  async function registerDevice(payload) {
    const input = normalizeRegistration(payload);
    const reference = devices.doc(input.deviceId);
    const ownerReference = tokenOwners.doc(input.tokenHash);
    let rotated = false;
    let migratedLegacy = false;

    await firestore.runTransaction(async (transaction) => {
      const snapshot = await transaction.get(reference);
      const ownerSnapshot = await transaction.get(ownerReference);
      const existing = snapshot.exists ? snapshot.data() : null;
      if (existing && !verifyHash(input.deviceSecret, existing.secretHash)) {
        throw new Error("UNAUTHORIZED_DEVICE");
      }
      const previousDeviceId = ownerSnapshot.exists ? String(ownerSnapshot.get("deviceId") || "") : "";
      const previousReference = previousDeviceId && previousDeviceId !== input.deviceId
        ? devices.doc(previousDeviceId)
        : null;
      const previousSnapshot = previousReference ? await transaction.get(previousReference) : null;
      const oldOwnerReference = existing?.tokenHash && existing.tokenHash !== input.tokenHash
        ? tokenOwners.doc(existing.tokenHash)
        : null;
      const oldOwnerSnapshot = oldOwnerReference ? await transaction.get(oldOwnerReference) : null;
      const timestamp = serverTimestamp();
      const previous = previousSnapshot?.exists ? previousSnapshot.data() : null;
      const inheritedKeywords = previous ? String(previous.keywords || "") : input.keywords;
      const inheritedActive = previous ? previous.active === true : input.active;

      if (previousReference && previous) {
        migratedLegacy = previous.legacy === true;
        transaction.update(previousReference, {
          active: false,
          token: null,
          claimedBy: input.deviceId,
          claimedAt: timestamp,
          updatedAt: timestamp
        });
      }

      if (!snapshot.exists) {
        transaction.create(reference, {
          schemaVersion: 1,
          deviceId: input.deviceId,
          secretHash: hash(input.deviceSecret),
          token: input.token,
          tokenHash: input.tokenHash,
          keywords: inheritedKeywords,
          userAgent: input.userAgent,
          active: inheritedActive,
          legacy: false,
          migratedFromLegacy: migratedLegacy,
          registeredAt: timestamp,
          updatedAt: timestamp,
          tokenUpdatedAt: timestamp,
          tokenRotationCount: 0,
          lastFailureCode: null,
          lastFailureAt: null
        });
      } else {
        rotated = existing.tokenHash !== input.tokenHash;
        transaction.update(reference, {
          token: input.token,
          tokenHash: input.tokenHash,
          previousTokenHash: rotated ? existing.tokenHash || null : existing.previousTokenHash || null,
          keywords: migratedLegacy ? inheritedKeywords : input.keywords,
          userAgent: input.userAgent,
          active: migratedLegacy ? inheritedActive : input.active,
          legacy: false,
          migratedFromLegacy: existing.migratedFromLegacy === true || migratedLegacy,
          updatedAt: timestamp,
          tokenUpdatedAt: rotated ? timestamp : existing.tokenUpdatedAt || timestamp,
          tokenRotationCount: Number(existing.tokenRotationCount || 0) + (rotated ? 1 : 0),
          lastFailureCode: null,
          lastFailureAt: null
        });
      }
      transaction.set(ownerReference, {
        tokenHash: input.tokenHash,
        deviceId: input.deviceId,
        legacy: false,
        updatedAt: timestamp
      });
      if (oldOwnerReference && oldOwnerSnapshot?.exists && oldOwnerSnapshot.get("deviceId") === input.deviceId) {
        transaction.delete(oldOwnerReference);
      }
    });

    const saved = await reference.get();
    return { ...publicStatus(saved.data()), rotated, migratedLegacy };
  }

  async function setDeviceActive(payload) {
    const device = await authenticatedDevice(payload);
    await device.reference.update({
      active: payload.active === true,
      keywords: boundedString(payload.keywords ?? device.data.keywords, 500, "INVALID_KEYWORDS"),
      updatedAt: serverTimestamp()
    });
    const saved = await device.reference.get();
    return publicStatus(saved.data());
  }

  async function getStatus(payload) {
    const device = await authenticatedDevice(payload);
    return publicStatus(device.data);
  }

  async function acknowledge(payload) {
    const input = normalizeReceipt(payload);
    const device = await authenticatedDevice(input);
    const timestamp = serverTimestamp();
    const receipt = device.reference.collection("receipts").doc(input.eventId);
    const receiptUpdate = {
      eventId: input.eventId,
      type: input.type,
      [`${input.phase}At`]: timestamp,
      updatedAt: timestamp
    };
    await receipt.set(receiptUpdate, { merge: true });

    const deviceField = input.type === "heartbeat" && input.phase === "received"
      ? "lastHeartbeatReceivedAt"
      : input.phase === "received"
        ? "lastNotificationReceivedAt"
        : input.phase === "shown"
          ? "lastShownAt"
          : "lastClickedAt";
    await device.reference.update({ [deviceField]: timestamp, updatedAt: timestamp });
    return { eventId: input.eventId, phase: input.phase, accepted: true };
  }

  async function markSendFailure(reference, error) {
    const unregistered = isUnregisteredError(error);
    await reference.update({
      active: unregistered ? false : true,
      token: unregistered ? null : (await reference.get()).data()?.token || null,
      lastFailureCode: String(error?.code || error?.message || "FCM_SEND_FAILED").slice(0, 200),
      lastFailureAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });
    return unregistered;
  }

  async function sendToDevice(reference, data, kind) {
    const snapshot = await reference.get();
    const device = snapshot.data();
    if (!snapshot.exists || device.active !== true || !device.token) throw new Error("INACTIVE_DEVICE");
    try {
      const messageId = await messaging.send({
        token: device.token,
        data,
        webpush: {
          headers: { TTL: kind === "heartbeat" ? "21600" : "86400", Urgency: "normal" },
          fcmOptions: { link: data.url }
        }
      });
      const acceptedField = kind === "heartbeat" ? "lastHeartbeatAcceptedAt" : "lastNotificationAcceptedAt";
      await reference.update({
        [acceptedField]: serverTimestamp(),
        lastFailureCode: null,
        lastFailureAt: null,
        updatedAt: serverTimestamp()
      });
      await reference.collection("receipts").doc(data.eventId).set({
        eventId: data.eventId,
        type: kind,
        acceptedAt: serverTimestamp(),
        messageId,
        updatedAt: serverTimestamp()
      }, { merge: true });
      return { accepted: true, eventId: data.eventId, messageId };
    } catch (error) {
      const unregistered = await markSendFailure(reference, error);
      logger.warn("Notification send failed", {
        deviceId: reference.id,
        code: String(error?.code || "FCM_SEND_FAILED"),
        unregistered
      });
      if (unregistered) throw new Error("FCM_TOKEN_UNREGISTERED");
      throw error;
    }
  }

  async function sendSelfTest(payload) {
    const device = await authenticatedDevice(payload);
    const latestPost = await latestPostProvider();
    const eventId = `selftest:${createEventId()}`;
    const result = await sendToDevice(device.reference, {
      eventId,
      type: "self-test",
      title: "테스트 알림 · 해피휴게더",
      body: latestPost.title,
      url: latestPost.link,
      icon: "https://somyun.github.io/BTCwebapp/icon-192.png"
    }, "self-test");
    return { ...result, postId: latestPost.id, postTitle: latestPost.title };
  }

  async function sendBoardPostAll(post) {
    const snapshot = await devices.get();
    const eventId = `board:${post.id}:v1`;
    const results = {
      scanned: snapshot.docs.length,
      matched: 0,
      accepted: 0,
      inactive: 0,
      keywordSkipped: 0,
      deduplicated: 0,
      failed: 0
    };
    for (const deviceSnapshot of snapshot.docs) {
      const device = deviceSnapshot.data();
      if (device.active !== true || !device.token) {
        results.inactive += 1;
        continue;
      }
      if (!matchesKeywords(post.title, device.keywords)) {
        results.keywordSkipped += 1;
        continue;
      }
      results.matched += 1;
      const receipt = await deviceSnapshot.ref.collection("receipts").doc(eventId).get();
      if (receipt.exists && receipt.data()?.acceptedAt) {
        results.deduplicated += 1;
        continue;
      }
      try {
        await sendToDevice(deviceSnapshot.ref, {
          eventId,
          type: "board-alert",
          title: "해피휴게더",
          body: post.title,
          url: post.link,
          icon: "https://somyun.github.io/BTCwebapp/icon-192.png"
        }, "board-alert");
        results.accepted += 1;
      } catch (_) {
        results.failed += 1;
      }
    }
    return { ...results, eventId, postId: post.id, postTitle: post.title };
  }

  async function previewLatestBoardPostTargets() {
    const boardPosts = await boardPostsProvider();
    if (!Array.isArray(boardPosts) || boardPosts.length === 0) throw new Error("HUMETRO_INVALID_BOARD_POSTS");
    const latestPost = [...boardPosts].sort((left, right) => left.id - right.id).at(-1);
    const snapshot = await devices.get();
    const result = { scanned: snapshot.docs.length, active: 0, matched: 0, inactive: 0, keywordSkipped: 0 };
    for (const deviceSnapshot of snapshot.docs) {
      const device = deviceSnapshot.data();
      if (device.active !== true || !device.token) {
        result.inactive += 1;
      } else {
        result.active += 1;
        if (matchesKeywords(latestPost.title, device.keywords)) result.matched += 1;
        else result.keywordSkipped += 1;
      }
    }
    return { ...result, postId: latestPost.id, postTitle: latestPost.title, postLink: latestPost.link };
  }

  async function initializeBoardPostBaseline() {
    const boardPosts = await boardPostsProvider();
    if (!Array.isArray(boardPosts) || boardPosts.length === 0) throw new Error("HUMETRO_INVALID_BOARD_POSTS");
    const latestPost = [...boardPosts].sort((left, right) => left.id - right.id).at(-1);
    const timestamp = serverTimestamp();
    await notificationState.set({
      lastPostId: latestPost.id,
      latestPostId: latestPost.id,
      latestPostTitle: latestPost.title,
      latestPostLink: latestPost.link,
      initializedAt: timestamp,
      lastCheckedAt: timestamp,
      updatedAt: timestamp
    }, { merge: true });
    return { initialized: true, postId: latestPost.id, postTitle: latestPost.title, postLink: latestPost.link };
  }

  async function sendLatestBoardPostIfNew() {
    const boardPosts = await boardPostsProvider();
    if (!Array.isArray(boardPosts) || boardPosts.length === 0) throw new Error("HUMETRO_INVALID_BOARD_POSTS");
    const sortedPosts = [...boardPosts].sort((left, right) => left.id - right.id);
    const latestPost = sortedPosts.at(-1);
    const stateSnapshot = await notificationState.get();
    const state = stateSnapshot.exists ? stateSnapshot.data() : {};
    const lastPostId = Number(state.lastPostId || 0);
    const timestamp = serverTimestamp();
    const latestState = {
      lastCheckedAt: timestamp,
      latestPostId: latestPost.id,
      latestPostTitle: latestPost.title,
      latestPostLink: latestPost.link,
      updatedAt: timestamp
    };

    if (!Number.isSafeInteger(lastPostId) || lastPostId <= 0) {
      await notificationState.set({
        ...latestState,
        lastPostId: latestPost.id,
        initializedAt: timestamp
      }, { merge: true });
      return { initialized: true, newPost: false, postId: latestPost.id, accepted: 0 };
    }

    if (latestPost.id <= lastPostId) {
      await notificationState.set(latestState, { merge: true });
      return { initialized: false, newPost: false, postId: latestPost.id, accepted: 0 };
    }

    const newPosts = sortedPosts.filter((post) => post.id > lastPostId);
    const dispatchResults = [];
    let completedPostId = lastPostId;
    for (const post of newPosts) {
      const result = await sendBoardPostAll(post);
      dispatchResults.push(result);
      if (result.failed > 0) break;
      completedPostId = post.id;
    }
    const result = dispatchResults.at(-1);
    const dispatchState = {
      ...latestState,
      lastDispatchedAt: timestamp,
      lastDispatchResults: dispatchResults
    };
    if (result.failed === 0) {
      dispatchState.lastPostId = completedPostId;
      dispatchState.retryPostId = null;
    } else {
      dispatchState.retryPostId = result.postId;
    }
    await notificationState.set(dispatchState, { merge: true });
    return {
      initialized: false,
      newPost: true,
      dispatchCount: dispatchResults.length,
      retryPending: result.failed > 0,
      ...result
    };
  }

  async function sendHeartbeatAll() {
    const snapshot = await devices.get();
    const results = { scanned: snapshot.docs.length, accepted: 0, inactive: 0, failed: 0 };
    for (const deviceSnapshot of snapshot.docs) {
      const device = deviceSnapshot.data();
      if (device.active !== true || !device.token || device.legacy === true) {
        results.inactive += 1;
        continue;
      }
      const eventId = `heartbeat:${now().toISOString().slice(0, 10)}:${createEventId()}`;
      try {
        await sendToDevice(deviceSnapshot.ref, {
          eventId,
          type: "heartbeat",
          title: "",
          body: "",
          url: "https://somyun.github.io/BTCwebapp/notification-settings.html",
          icon: "https://somyun.github.io/BTCwebapp/icon-192.png"
        }, "heartbeat");
        results.accepted += 1;
      } catch (_) {
        results.failed += 1;
      }
    }
    return results;
  }

  return {
    acknowledge,
    getStatus,
    initializeBoardPostBaseline,
    previewLatestBoardPostTargets,
    registerDevice,
    sendLatestBoardPostIfNew,
    sendHeartbeatAll,
    sendSelfTest,
    setDeviceActive
  };
}

module.exports = {
  createNotificationService,
  hash,
  isUnregisteredError,
  matchesKeywords,
  normalizeCredentials,
  normalizeReceipt,
  normalizeRegistration,
  publicStatus
};
