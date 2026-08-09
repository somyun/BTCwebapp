"use strict";

const { createHash } = require("node:crypto");
const { createReadStream, readdirSync, readFileSync, statSync } = require("node:fs");
const path = require("node:path");

const PROJECT_BUCKETS = Object.freeze({
  "btcwebapp-551bd": "btcwebapp-551bd.firebasestorage.app",
  "btcwebapp-test": "btcwebapp-test.firebasestorage.app"
});
const OBJECT_PREFIX = "cad/hopo/";
const CLI_ROOT = path.join(process.env.APPDATA || "", "npm", "node_modules", "firebase-tools", "lib");
const auth = require(path.join(CLI_ROOT, "auth.js"));
const { Client, setRefreshToken } = require(path.join(CLI_ROOT, "apiv2.js"));
const firebaseStorage = require(path.join(CLI_ROOT, "gcp", "storage.js"));

function filesBelow(directory, base = directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(directory, entry.name);
    return entry.isDirectory()
      ? filesBelow(fullPath, base)
      : [{ fullPath, relativePath: path.relative(base, fullPath).replaceAll(path.sep, "/") }];
  });
}

function md5Base64(file) {
  return createHash("md5").update(readFileSync(file)).digest("base64");
}

async function configureCliAuth() {
  const account = auth.getGlobalDefaultAccount();
  if (!account?.tokens?.refresh_token) throw new Error("FIREBASE_CLI_LOGIN_REQUIRED");
  setRefreshToken(account.tokens.refresh_token);
}

async function listObjects(client, bucket) {
  const objects = [];
  let pageToken;
  do {
    const response = await client.get(`/storage/v1/b/${bucket}/o`, {
      queryParams: {
        prefix: OBJECT_PREFIX,
        ...(pageToken ? { pageToken } : {})
      }
    });
    objects.push(...(response.body.items || []));
    pageToken = response.body.nextPageToken;
  } while (pageToken);
  return objects;
}

async function configureCors(client, bucket) {
  await client.patch(`/storage/v1/b/${bucket}`, {
    cors: [{
      origin: ["https://somyun.github.io"],
      method: ["GET"],
      responseHeader: ["Content-Type", "ETag", "Cache-Control", "x-goog-generation"],
      maxAgeSeconds: 3600
    }]
  }, { queryParams: { updateMask: "cors" } });
}

async function uploadFile(client, bucket, entry) {
  const objectName = `${OBJECT_PREFIX}${entry.relativePath}`;
  await client.request({
    method: "POST",
    path: `/upload/storage/v1/b/${bucket}/o`,
    queryParams: { uploadType: "media", name: objectName },
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: createReadStream(entry.fullPath),
    skipLog: { reqBody: true, resBody: true }
  });
  await client.patch(`/storage/v1/b/${bucket}/o/${encodeURIComponent(objectName)}`, {
    contentType: "application/json; charset=utf-8",
    cacheControl: "private, max-age=3600, no-transform"
  }, { queryParams: { updateMask: "contentType,cacheControl" } });
  const stored = await client.get(`/storage/v1/b/${bucket}/o/${encodeURIComponent(objectName)}`);
  if (stored.body.md5Hash !== md5Base64(entry.fullPath) ||
      Number(stored.body.size) !== statSync(entry.fullPath).size) {
    throw new Error(`UPLOAD_VERIFICATION_FAILED:${objectName}`);
  }
  return objectName;
}

async function syncProject(projectId, sourceDirectory) {
  const expectedBucket = PROJECT_BUCKETS[projectId];
  if (!expectedBucket) throw new Error(`PROJECT_NOT_ALLOWED:${projectId}`);
  const defaultBucket = await firebaseStorage.getDefaultBucket(projectId);
  if (defaultBucket !== expectedBucket) {
    throw new Error(`DEFAULT_BUCKET_MISMATCH:${projectId}:${defaultBucket}`);
  }
  const client = new Client({ urlPrefix: "https://storage.googleapis.com" });
  await configureCors(client, expectedBucket);
  const entries = filesBelow(sourceDirectory);
  const uploaded = new Set();
  for (const entry of entries) uploaded.add(await uploadFile(client, expectedBucket, entry));
  const storedObjects = await listObjects(client, expectedBucket);
  for (const object of storedObjects) {
    if (!uploaded.has(object.name)) {
      await client.delete(`/storage/v1/b/${expectedBucket}/o/${encodeURIComponent(object.name)}`);
    }
  }
  const verified = await listObjects(client, expectedBucket);
  if (verified.length !== entries.length) throw new Error(`OBJECT_COUNT_MISMATCH:${projectId}`);
  return {
    projectId,
    bucket: expectedBucket,
    objectCount: verified.length,
    totalBytes: entries.reduce((sum, entry) => sum + statSync(entry.fullPath).size, 0)
  };
}

async function main() {
  const sourceDirectory = path.resolve(process.argv[2] || path.join(__dirname, "..", "cad-data", "hopo"));
  const projects = process.argv.slice(3);
  if (!projects.length) throw new Error("PROJECT_ID_REQUIRED");
  await configureCliAuth();
  const results = [];
  for (const projectId of projects) results.push(await syncProject(projectId, sourceDirectory));
  process.stdout.write(`${JSON.stringify({ success: true, results }, null, 2)}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
