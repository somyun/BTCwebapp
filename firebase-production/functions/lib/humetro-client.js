"use strict";

const HUMETRO_HOST = "www.humetro.busan.kr";
const HUMETRO_LOGIN_URL = "https://www.humetro.busan.kr/homepage/default/member/page/loginProcEvent.do";
const HUMETRO_BOARD_URL = "https://www.humetro.busan.kr/homepage/default/board/listEvent.do?conf_no=151&menu_no=1001060402";

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

function decodeHtmlEntities(value) {
  return String(value || "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, "\"")
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_match, code) => String.fromCodePoint(Number(code)));
}

function parseBoardHtml(html) {
  const tbody = String(html || "").match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/i);
  if (!tbody) throw new Error("HUMETRO_BOARD_PARSE_FAILED");
  const posts = [];
  for (const row of tbody[1].split(/<\/tr>/i)) {
    const numberCell = row.match(/<th[^>]*>([\s\S]*?)<\/th>/i);
    if (!numberCell) continue;
    const idText = numberCell[1].replace(/<[^>]+>/g, "").trim();
    if (!/^\d+$/.test(idText)) continue;
    const subject = row.match(/<td[^>]*class=["']?subject["']?[^>]*>([\s\S]*?)<\/td>/i);
    if (!subject) continue;
    const anchor = subject[1].match(/<a[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i);
    if (!anchor) continue;
    const link = new URL(decodeHtmlEntities(anchor[1]), `https://${HUMETRO_HOST}`).href;
    const title = decodeHtmlEntities(anchor[2].replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim();
    posts.push({ id: Number(idText), title, link });
  }
  return normalizeBoardPosts({ success: true, data: posts });
}

function responseCookies(headers) {
  const values = typeof headers?.getSetCookie === "function"
    ? headers.getSetCookie()
    : [headers?.get?.("set-cookie") || ""];
  return values
    .flatMap((value) => String(value).split(/,(?=\s*[^;,=]+=[^;,]+)/))
    .map((value) => value.split(";")[0].trim())
    .filter(Boolean)
    .join("; ");
}

async function responseText(response, fallbackCharset = "euc-kr") {
  const bytes = await response.arrayBuffer();
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch (_) {
    return new TextDecoder(fallbackCharset).decode(bytes);
  }
}

function createHumetroBoardPostsProvider({
  fetchImpl = globalThis.fetch,
  credentialsProvider,
  timeoutMs = 20000
} = {}) {
  if (typeof fetchImpl !== "function" || typeof credentialsProvider !== "function") {
    throw new Error("INVALID_HUMETRO_FETCH_DEPENDENCY");
  }

  return async function humetroBoardPostsProvider() {
    const credentials = await credentialsProvider();
    const userId = String(credentials?.userId || "").trim();
    const password = String(credentials?.password || "");
    if (!userId || !password || userId.length > 200 || password.length > 300) {
      throw new Error("HUMETRO_CREDENTIALS_UNAVAILABLE");
    }

    const loginBody = new URLSearchParams({ RETURNURL: "", userID: userId, password });
    let loginResponse;
    try {
      loginResponse = await fetchImpl(HUMETRO_LOGIN_URL, {
        method: "POST",
        redirect: "manual",
        headers: {
          Accept: "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
          "Accept-Language": "ko,en;q=0.9,en-US;q=0.8",
          "Cache-Control": "max-age=0",
          "Content-Type": "application/x-www-form-urlencoded",
          DNT: "1",
          Origin: "https://www.humetro.busan.kr",
          Referer: "https://www.humetro.busan.kr/event.do",
          "Sec-Ch-Ua": '"Not(A:Brand";v="8", "Chromium";v="144", "Microsoft Edge";v="144"',
          "Sec-Ch-Ua-Mobile": "?0",
          "Sec-Ch-Ua-Platform": '"Windows"',
          "Sec-Fetch-Dest": "document",
          "Sec-Fetch-Mode": "navigate",
          "Sec-Fetch-Site": "same-origin",
          "Sec-Fetch-User": "?1",
          "Upgrade-Insecure-Requests": "1",
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0"
        },
        body: loginBody.toString(),
        signal: AbortSignal.timeout(timeoutMs)
      });
    } catch (error) {
      throw new Error(`HUMETRO_LOGIN_FETCH_FAILED:${String(error?.message || error).slice(0, 120)}`);
    }
    if (loginResponse.status < 200 || loginResponse.status >= 400) {
      throw new Error(`HUMETRO_LOGIN_HTTP_${loginResponse.status}`);
    }
    const cookie = responseCookies(loginResponse.headers);
    if (!cookie) throw new Error("HUMETRO_LOGIN_COOKIE_MISSING");

    let boardResponse;
    try {
      boardResponse = await fetchImpl(HUMETRO_BOARD_URL, {
        method: "POST",
        redirect: "follow",
        headers: {
          Accept: "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
          "Accept-Language": "ko,en;q=0.9,en-US;q=0.8",
          "Cache-Control": "max-age=0",
          "Content-Type": "application/x-www-form-urlencoded",
          Cookie: cookie,
          Origin: "https://www.humetro.busan.kr",
          Referer: HUMETRO_LOGIN_URL,
          "Upgrade-Insecure-Requests": "1",
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36"
        },
        body: new URLSearchParams({ RETURNURL: "null" }).toString(),
        signal: AbortSignal.timeout(timeoutMs)
      });
    } catch (error) {
      throw new Error(`HUMETRO_BOARD_FETCH_FAILED:${String(error?.message || error).slice(0, 120)}`);
    }
    if (!boardResponse.ok) throw new Error(`HUMETRO_BOARD_HTTP_${boardResponse.status}`);
    return parseBoardHtml(await responseText(boardResponse));
  };
}

function createHumetroLatestPostProvider(options = {}) {
  const boardPostsProvider = createHumetroBoardPostsProvider(options);
  return async function humetroLatestPostProvider() {
    const posts = await boardPostsProvider();
    return posts.at(-1);
  };
}

module.exports = {
  HUMETRO_BOARD_URL,
  HUMETRO_HOST,
  HUMETRO_LOGIN_URL,
  createHumetroBoardPostsProvider,
  createHumetroLatestPostProvider,
  normalizeBoardPosts,
  normalizeLatestPost,
  parseBoardHtml,
  responseCookies
};
