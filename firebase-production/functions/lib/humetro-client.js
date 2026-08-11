"use strict";

const PRODUCTION_GAS_API_URL = "https://script.google.com/macros/s/AKfycbzuWS4Q5kTzDRH4IBpeXBa69KngElRdArtTCzTV0NDQsB3y4oABBIzrTLuPOZH5KOPP/exec";
const HUMETRO_HOST = "www.humetro.busan.kr";

function normalizeLatestPost(payload = {}) {
  const source = payload?.success === true ? payload.data : null;
  const id = Number(source?.id);
  const title = String(source?.title || "").replace(/\s+/g, " ").trim();
  let link;
  try {
    link = new URL(String(source?.link || ""));
  } catch (_) {
    throw new Error("HUMETRO_INVALID_LATEST_POST");
  }

  if (!Number.isSafeInteger(id) || id <= 0 || !title || title.length > 300) {
    throw new Error("HUMETRO_INVALID_LATEST_POST");
  }
  if (link.protocol !== "https:" || link.hostname !== HUMETRO_HOST ||
      !link.pathname.startsWith("/homepage/default/board/")) {
    throw new Error("HUMETRO_INVALID_LATEST_POST_LINK");
  }

  return { id, title, link: link.href };
}

function normalizeBoardPosts(payload = {}) {
  const source = payload?.success === true ? payload.data : null;
  if (!Array.isArray(source) || source.length === 0 || source.length > 100) {
    throw new Error("HUMETRO_INVALID_BOARD_POSTS");
  }
  const posts = source.map((post) => normalizeLatestPost({ success: true, data: post }));
  return [...new Map(posts.map((post) => [post.id, post])).values()]
    .sort((left, right) => left.id - right.id);
}

function createHumetroProvider({
  action,
  normalize,
  fetchImpl = globalThis.fetch,
  gasApiUrl = PRODUCTION_GAS_API_URL,
  bridgeTokenProvider,
  timeoutMs = 20000
} = {}) {
  if (typeof fetchImpl !== "function" || typeof bridgeTokenProvider !== "function") {
    throw new Error("INVALID_HUMETRO_FETCH_DEPENDENCY");
  }
  const endpoint = new URL(gasApiUrl);

  return async function humetroProvider() {
    const bridgeToken = String(await bridgeTokenProvider()).trim();
    if (!/^[A-Za-z0-9_-]{32,160}$/.test(bridgeToken)) throw new Error("HUMETRO_BRIDGE_TOKEN_UNAVAILABLE");
    let response;
    try {
      response = await fetchImpl(endpoint, {
        method: "POST",
        redirect: "follow",
        headers: { Accept: "application/json", "Content-Type": "application/json" },
        body: JSON.stringify({ action, bridgeToken }),
        signal: AbortSignal.timeout(timeoutMs)
      });
    } catch (error) {
      throw new Error(`HUMETRO_LATEST_POST_FETCH_FAILED:${String(error?.message || error).slice(0, 120)}`);
    }
    if (!response.ok) throw new Error(`HUMETRO_LATEST_POST_HTTP_${response.status}`);

    let payload;
    try {
      payload = await response.json();
    } catch (_) {
      throw new Error("HUMETRO_LATEST_POST_INVALID_JSON");
    }
    if (payload?.error) throw new Error(`HUMETRO_LATEST_POST_SOURCE_ERROR:${String(payload.error).slice(0, 120)}`);
    return normalize(payload);
  };
}

function createHumetroLatestPostProvider({
  fetchImpl = globalThis.fetch,
  gasApiUrl = PRODUCTION_GAS_API_URL,
  bridgeTokenProvider,
  timeoutMs = 20000
} = {}) {
  return createHumetroProvider({
    action: "getLatestBoardPostForNotificationTest",
    normalize: normalizeLatestPost,
    fetchImpl,
    gasApiUrl,
    bridgeTokenProvider,
    timeoutMs
  });
}

function createHumetroBoardPostsProvider(options = {}) {
  return createHumetroProvider({
    ...options,
    action: "getBoardPostsForNotificationDispatch",
    normalize: normalizeBoardPosts
  });
}

module.exports = {
  PRODUCTION_GAS_API_URL,
  createHumetroBoardPostsProvider,
  createHumetroLatestPostProvider,
  normalizeBoardPosts,
  normalizeLatestPost
};
