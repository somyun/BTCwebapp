"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const {
  createHumetroBoardPostsProvider,
  createHumetroLatestPostProvider,
  normalizeBoardPosts,
  normalizeLatestPost
} = require("../lib/humetro-client");

const sourcePayload = {
  success: true,
  data: {
    id: 6042,
    title: "  최신   경조사 게시글  ",
    link: "https://www.humetro.busan.kr/homepage/default/board/viewEvent.do?board_no=6042"
  }
};

test("normalizes the latest post returned by the isolated GAS", () => {
  assert.deepEqual(normalizeLatestPost(sourcePayload), {
    id: 6042,
    title: "최신 경조사 게시글",
    link: sourcePayload.data.link
  });
});

test("rejects a latest-post link outside the Happy Hugether host", () => {
  assert.throws(() => normalizeLatestPost({
    ...sourcePayload,
    data: { ...sourcePayload.data, link: "https://example.com/fake" }
  }), { message: "HUMETRO_INVALID_LATEST_POST_LINK" });
});

test("normalizes, sorts, and deduplicates recent board posts", () => {
  const older = { ...sourcePayload.data, id: 6041, title: "이전 게시글", link: sourcePayload.data.link.replace("6042", "6041") };
  assert.deepEqual(normalizeBoardPosts({ success: true, data: [sourcePayload.data, older, sourcePayload.data] }), [
    { id: 6041, title: "이전 게시글", link: older.link },
    { id: 6042, title: "최신 경조사 게시글", link: sourcePayload.data.link }
  ]);
});

test("requests the keyword-independent latest-post action from test GAS", async () => {
  let requestedUrl = "";
  let requestedOptions;
  const provider = createHumetroLatestPostProvider({
    gasApiUrl: "https://script.google.com/macros/s/test/exec",
    bridgeTokenProvider: async () => "b".repeat(64),
    fetchImpl: async (url, options) => {
      requestedUrl = String(url);
      requestedOptions = options;
      return { ok: true, async json() { return sourcePayload; } };
    }
  });
  const post = await provider();
  assert.equal(new URL(requestedUrl).search, "");
  assert.equal(requestedOptions.method, "POST");
  assert.deepEqual(JSON.parse(requestedOptions.body), {
    action: "getLatestBoardPostForNotificationTest",
    bridgeToken: "b".repeat(64)
  });
  assert.equal(post.id, 6042);
});

test("requests the private recent-post list for scheduled dispatch", async () => {
  let requestedOptions;
  const provider = createHumetroBoardPostsProvider({
    gasApiUrl: "https://script.google.com/macros/s/test/exec",
    bridgeTokenProvider: async () => "b".repeat(64),
    fetchImpl: async (_url, options) => {
      requestedOptions = options;
      return { ok: true, async json() { return { success: true, data: [sourcePayload.data] }; } };
    }
  });
  const posts = await provider();
  assert.deepEqual(JSON.parse(requestedOptions.body), {
    action: "getBoardPostsForNotificationDispatch",
    bridgeToken: "b".repeat(64)
  });
  assert.equal(posts[0].id, 6042);
});

test("reports invalid GAS JSON without sending a notification", async () => {
  const provider = createHumetroLatestPostProvider({
    bridgeTokenProvider: async () => "b".repeat(64),
    fetchImpl: async () => ({
      ok: true,
      async json() { throw new SyntaxError("bad json"); }
    })
  });
  await assert.rejects(provider(), { message: "HUMETRO_LATEST_POST_INVALID_JSON" });
});

test("does not call public GAS when the private bridge token is unavailable", async () => {
  let called = false;
  const provider = createHumetroLatestPostProvider({
    bridgeTokenProvider: async () => "",
    fetchImpl: async () => {
      called = true;
      throw new Error("unexpected");
    }
  });
  await assert.rejects(provider(), { message: "HUMETRO_BRIDGE_TOKEN_UNAVAILABLE" });
  assert.equal(called, false);
});
