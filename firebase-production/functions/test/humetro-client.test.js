"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const {
  HUMETRO_BOARD_URL,
  HUMETRO_LOGIN_URL,
  createHumetroBoardPostsProvider,
  createHumetroLatestPostProvider,
  normalizeBoardPosts,
  normalizeLatestPost,
  parseBoardHtml
} = require("../lib/humetro-client");

const sourcePayload = {
  success: true,
  data: {
    id: 6042,
    title: "  최신   경조사 게시글  ",
    link: "https://www.humetro.busan.kr/homepage/default/board/viewEvent.do?board_no=6042"
  }
};

function html(posts = [sourcePayload.data]) {
  return `<table><tbody>${posts.map((post) => `<tr><th>${post.id}</th><td class="subject"><a href="${post.link}">${post.title}</a></td></tr>`).join("")}</tbody></table>`;
}

function response({ status = 200, body = "", cookies = [] } = {}) {
  return {
    status,
    ok: status >= 200 && status < 300,
    headers: { getSetCookie: () => cookies },
    async arrayBuffer() { return new TextEncoder().encode(body).buffer; }
  };
}

test("normalizes and validates a latest Happy Hugether post", () => {
  assert.deepEqual(normalizeLatestPost(sourcePayload), {
    id: 6042,
    title: "최신 경조사 게시글",
    link: sourcePayload.data.link
  });
  assert.throws(() => normalizeLatestPost({
    ...sourcePayload,
    data: { ...sourcePayload.data, link: "https://example.com/fake" }
  }), { message: "HUMETRO_INVALID_LATEST_POST_LINK" });
});

test("normalizes, sorts, and deduplicates recent posts", () => {
  const older = { ...sourcePayload.data, id: 6041, title: "이전 게시글", link: sourcePayload.data.link.replace("6042", "6041") };
  assert.deepEqual(normalizeBoardPosts({ success: true, data: [sourcePayload.data, older, sourcePayload.data] }), [
    { id: 6041, title: "이전 게시글", link: older.link },
    { id: 6042, title: "최신 경조사 게시글", link: sourcePayload.data.link }
  ]);
});

test("parses the production board table without a GAS bridge", () => {
  assert.deepEqual(parseBoardHtml(html()), [{
    id: 6042,
    title: "최신 경조사 게시글",
    link: sourcePayload.data.link
  }]);
});

test("logs in directly, preserves the session cookie, and reads recent posts", async () => {
  const calls = [];
  const provider = createHumetroBoardPostsProvider({
    credentialsProvider: async () => ({ userId: "employee", password: "secret" }),
    fetchImpl: async (url, options) => {
      calls.push({ url: String(url), options });
      if (String(url) === HUMETRO_LOGIN_URL) {
        return response({ status: 302, cookies: ["JSESSIONID=session123; Path=/; HttpOnly"] });
      }
      return response({ body: html() });
    }
  });
  const posts = await provider();
  assert.equal(calls.length, 2);
  assert.equal(calls[0].url, HUMETRO_LOGIN_URL);
  assert.match(calls[0].options.body, /userID=employee/);
  assert.equal(calls[1].url, HUMETRO_BOARD_URL);
  assert.equal(calls[1].options.headers.Cookie, "JSESSIONID=session123");
  assert.equal(posts[0].id, 6042);
});

test("latest provider returns the greatest post id", async () => {
  const older = { ...sourcePayload.data, id: 6041, title: "이전 게시글", link: sourcePayload.data.link.replace("6042", "6041") };
  let call = 0;
  const provider = createHumetroLatestPostProvider({
    credentialsProvider: async () => ({ userId: "employee", password: "secret" }),
    fetchImpl: async () => (++call === 1
      ? response({ status: 302, cookies: ["JSESSIONID=session123; Path=/"] })
      : response({ body: html([sourcePayload.data, older]) }))
  });
  assert.equal((await provider()).id, 6042);
});

test("credentials are required before any external request", async () => {
  let called = false;
  const provider = createHumetroBoardPostsProvider({
    credentialsProvider: async () => ({ userId: "", password: "" }),
    fetchImpl: async () => { called = true; }
  });
  await assert.rejects(provider(), { message: "HUMETRO_CREDENTIALS_UNAVAILABLE" });
  assert.equal(called, false);
});
