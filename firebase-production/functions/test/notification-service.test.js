"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const { createNotificationService, matchesKeywords } = require("../lib/notification-service");
const { FakeFirestore } = require("./helpers/fake-firestore");

const deviceId = `d_${"a".repeat(43)}`;
const deviceSecret = "s".repeat(64);
const registration = {
  deviceId,
  deviceSecret,
  token: `token-${"x".repeat(80)}`,
  keywords: "점검",
  userAgent: "Unit Test",
  active: true
};

function createMessaging() {
  return {
    sent: [],
    async send(message) {
      this.sent.push(structuredClone(message));
      return `projects/test/messages/${this.sent.length}`;
    }
  };
}

const latestPost = {
  id: 6042,
  title: "최신 경조사 게시글",
  link: "https://www.humetro.busan.kr/homepage/default/board/viewEvent.do?board_no=6042"
};

function createService({
  firestore = new FakeFirestore(),
  messaging = createMessaging(),
  latestPostProvider = async () => latestPost
} = {}) {
  return {
    firestore,
    messaging,
    service: createNotificationService({
      firestore,
      messaging,
      serverTimestamp: () => "SERVER_TIMESTAMP",
      latestPostProvider,
      now: () => new Date("2026-07-31T00:00:00.000Z"),
      createEventId: () => "123e4567-e89b-12d3-a456-426614174000",
      logger: { warn() {} }
    })
  };
}

test("registers one private device document and returns no credentials", async () => {
  const { service, firestore } = createService();
  const result = await service.registerDevice(registration);
  assert.equal(result.active, true);
  assert.equal(result.rotated, false);
  assert.equal("token" in result, false);
  assert.equal("secretHash" in result, false);
  const stored = firestore.documents.get(`notificationDevices/${deviceId}`);
  assert.equal(stored.token, registration.token);
  assert.notEqual(stored.secretHash, deviceSecret);
});

test("rotates the token in the same device document instead of appending a device", async () => {
  const { service, firestore } = createService();
  await service.registerDevice(registration);
  const nextToken = `token-${"y".repeat(80)}`;
  const result = await service.registerDevice({ ...registration, token: nextToken });
  assert.equal(result.rotated, true);
  assert.equal(firestore.documents.get(`notificationDevices/${deviceId}`).token, nextToken);
  const devicePaths = [...firestore.documents.keys()].filter((path) => /^notificationDevices\/[^/]+$/.test(path));
  assert.deepEqual(devicePaths, [`notificationDevices/${deviceId}`]);
});

test("rejects a caller that does not hold the device secret", async () => {
  const { service } = createService();
  await service.registerDevice(registration);
  await assert.rejects(service.getStatus({ ...registration, deviceSecret: "z".repeat(64) }), {
    message: "UNAUTHORIZED_DEVICE"
  });
});

test("sends a targeted self-test and records FCM acceptance", async () => {
  const { service, messaging, firestore } = createService();
  await service.registerDevice(registration);
  const result = await service.sendSelfTest(registration);
  assert.equal(result.accepted, true);
  assert.match(result.eventId, /^selftest:/);
  assert.equal(messaging.sent.length, 1);
  assert.equal(messaging.sent[0].token, registration.token);
  assert.equal(messaging.sent[0].data.type, "self-test");
  assert.equal(messaging.sent[0].data.title, "테스트 알림 · 해피휴게더");
  assert.equal(messaging.sent[0].data.body, latestPost.title);
  assert.equal(messaging.sent[0].data.url, latestPost.link);
  assert.equal(result.postId, latestPost.id);
  assert.equal(result.postTitle, latestPost.title);
  assert.equal(
    firestore.documents.get(`notificationDevices/${deviceId}/receipts/${result.eventId}`).messageId,
    "projects/test/messages/1"
  );
});

test("self-test fetches the latest Happy Hugether post regardless of device keywords", async () => {
  let calls = 0;
  const provider = async () => {
    calls += 1;
    return latestPost;
  };
  const { service, messaging } = createService({ latestPostProvider: provider });
  await service.registerDevice({ ...registration, keywords: "절대 일치하지 않을 키워드" });
  await service.sendSelfTest(registration);
  assert.equal(calls, 1);
  assert.equal(messaging.sent[0].data.body, latestPost.title);
});

test("matches comma-separated OR groups and space-separated AND terms", () => {
  assert.equal(matchesKeywords("경전철 전기 분야 경조사", "전기사업소, 경전철 전기"), true);
  assert.equal(matchesKeywords("경전철 신호 분야 경조사", "전기사업소, 경전철 전기"), false);
  assert.equal(matchesKeywords("모든 새 게시글", ""), true);
});

test("initial scheduled check establishes a baseline without sending an old post", async () => {
  const { service, messaging, firestore } = createService();
  await service.registerDevice({ ...registration, keywords: "경조사" });
  const result = await service.sendLatestBoardPostIfNew();
  assert.deepEqual(result, { initialized: true, newPost: false, postId: latestPost.id, accepted: 0 });
  assert.equal(messaging.sent.length, 0);
  assert.equal(firestore.documents.get("systemConfig/happyHugetherNotifications").lastPostId, latestPost.id);
});

test("scheduled check sends a new post to active keyword-matched devices only", async () => {
  let currentPost = latestPost;
  const { service, messaging, firestore } = createService({ latestPostProvider: async () => currentPost });
  await service.registerDevice({ ...registration, keywords: "경조사" });
  await service.registerDevice({
    ...registration,
    deviceId: `d_${"b".repeat(43)}`,
    deviceSecret: "q".repeat(64),
    token: `token-${"y".repeat(80)}`,
    keywords: "전기사업소"
  });
  await service.registerDevice({
    ...registration,
    deviceId: `d_${"c".repeat(43)}`,
    deviceSecret: "r".repeat(64),
    token: `token-${"z".repeat(80)}`,
    active: false
  });
  await service.sendLatestBoardPostIfNew();
  currentPost = { ...latestPost, id: latestPost.id + 1, title: "새 경조사 게시글" };

  const result = await service.sendLatestBoardPostIfNew();
  assert.equal(result.newPost, true);
  assert.equal(result.scanned, 3);
  assert.equal(result.matched, 1);
  assert.equal(result.accepted, 1);
  assert.equal(result.keywordSkipped, 1);
  assert.equal(result.inactive, 1);
  assert.equal(messaging.sent.length, 1);
  assert.equal(messaging.sent[0].data.type, "board-alert");
  assert.equal(messaging.sent[0].data.eventId, `board:${currentPost.id}:v1`);
  assert.equal(firestore.documents.get("systemConfig/happyHugetherNotifications").lastPostId, currentPost.id);
});

test("scheduled dispatch skips a deterministic event already accepted for a device", async () => {
  const firestore = new FakeFirestore({
    "systemConfig/happyHugetherNotifications": { lastPostId: latestPost.id - 1 }
  });
  const { service, messaging } = createService({ firestore });
  await service.registerDevice({ ...registration, keywords: "경조사" });
  firestore.documents.set(
    `notificationDevices/${deviceId}/receipts/board:${latestPost.id}:v1`,
    { acceptedAt: "SERVER_TIMESTAMP" }
  );
  const result = await service.sendLatestBoardPostIfNew();
  assert.equal(result.deduplicated, 1);
  assert.equal(result.accepted, 0);
  assert.equal(messaging.sent.length, 0);
});

test("scheduled dispatch preserves every new post returned by the bridge", async () => {
  const posts = [
    { ...latestPost, id: latestPost.id + 1, title: "첫 번째 새 경조사" },
    { ...latestPost, id: latestPost.id + 2, title: "두 번째 새 경조사" }
  ];
  const firestore = new FakeFirestore({
    "systemConfig/happyHugetherNotifications": { lastPostId: latestPost.id }
  });
  const messaging = createMessaging();
  const service = createNotificationService({
    firestore,
    messaging,
    serverTimestamp: () => "SERVER_TIMESTAMP",
    latestPostProvider: async () => posts.at(-1),
    boardPostsProvider: async () => posts,
    now: () => new Date("2026-07-31T00:00:00.000Z"),
    createEventId: () => "123e4567-e89b-12d3-a456-426614174000",
    logger: { warn() {} }
  });
  await service.registerDevice({ ...registration, keywords: "경조사" });
  const result = await service.sendLatestBoardPostIfNew();
  assert.equal(result.dispatchCount, 2);
  assert.deepEqual(messaging.sent.map((message) => message.data.body), posts.map((post) => post.title));
  assert.equal(firestore.documents.get("systemConfig/happyHugetherNotifications").lastPostId, posts.at(-1).id);
});

test("deactivates and clears a token rejected as UNREGISTERED", async () => {
  const messaging = {
    async send() {
      const error = new Error("Requested entity was not found: UNREGISTERED");
      error.code = "messaging/registration-token-not-registered";
      throw error;
    }
  };
  const { service, firestore } = createService({ messaging });
  await service.registerDevice(registration);
  await assert.rejects(service.sendSelfTest(registration), { message: "FCM_TOKEN_UNREGISTERED" });
  const stored = firestore.documents.get(`notificationDevices/${deviceId}`);
  assert.equal(stored.active, false);
  assert.equal(stored.token, null);
});

test("daily heartbeat targets active devices only", async () => {
  const { service, messaging } = createService();
  await service.registerDevice(registration);
  const inactiveId = `d_${"b".repeat(43)}`;
  await service.registerDevice({ ...registration, deviceId: inactiveId, deviceSecret: "q".repeat(64), active: false });
  const result = await service.sendHeartbeatAll();
  assert.deepEqual(result, { scanned: 2, accepted: 1, inactive: 1, failed: 0 });
  assert.equal(messaging.sent[0].data.type, "heartbeat");
});

test("receipt updates both the event phase and device health timestamp", async () => {
  const { service, firestore } = createService();
  await service.registerDevice(registration);
  const receipt = await service.acknowledge({
    deviceId,
    deviceSecret,
    eventId: "selftest:event-1234",
    type: "self-test",
    phase: "received"
  });
  assert.equal(receipt.accepted, true);
  assert.equal(firestore.documents.get(`notificationDevices/${deviceId}`).lastNotificationReceivedAt, "SERVER_TIMESTAMP");
  assert.equal(
    firestore.documents.get(`notificationDevices/${deviceId}/receipts/selftest:event-1234`).receivedAt,
    "SERVER_TIMESTAMP"
  );
});
