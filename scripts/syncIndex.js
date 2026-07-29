// scripts/syncIndex.js
// Clears the Azure AI Search index and re-indexes all current PDFs from SharePoint.
// Use this whenever files are added, removed, or replaced in SharePoint.
// Run with: node scripts/syncIndex.js

const fs   = require("fs");
const path = require("path");
const pdf  = require("pdf-parse");
const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");

// Load env vars from local.settings.json
const settings = JSON.parse(fs.readFileSync(path.join(__dirname, "../local.settings.json"), "utf8"));
Object.assign(process.env, settings.Values);

// ─────────────────────────────────────────────
// CONFIGURATION
// ─────────────────────────────────────────────

const SEARCH_SERVICE_NAME = process.env.SEARCH_SERVICE_NAME;
const SEARCH_API_KEY      = process.env.SEARCH_API_KEY;
const INDEX_NAME          = process.env.SEARCH_INDEX_NAME || "sharepoint-index";

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const SP_HOSTNAME  = "chimeratechpvtltd.sharepoint.com";
const SP_SITE_PATH = "/sites/Corporate-Policies";

const FOLDER_GROUP_MAP = {
  General_Policies:  ["2f996595-97dc-45dd-a85a-0d20759e1682"],
  Manager_Policies:  ["70205b81-db45-4fa4-92b7-7fb87a89c73f"],
  Sales_Policies:    ["0941d8fd-91c2-415e-8b74-e81cf00d08ea"],
};

// ─────────────────────────────────────────────
// AUTH + SHAREPOINT HELPERS (same as indexDocuments.js)
// ─────────────────────────────────────────────

async function getAccessToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default",
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error(`Auth failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function getSiteAndDrive(token) {
  const siteRes  = await fetch(`https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const siteData = await siteRes.json();
  if (siteData.error) throw new Error(`getSite failed: ${siteData.error.message}`);

  const driveRes  = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/drives`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const driveData = await driveRes.json();
  if (driveData.error) throw new Error(`getDrive failed: ${driveData.error.message}`);

  const drive = driveData.value.find(d => d.name === "Documents" || d.driveType === "documentLibrary") || driveData.value[0];
  return { siteId: siteData.id, driveId: drive.id };
}

async function listFilesInFolder(token, driveId, folderName) {
  const res  = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${folderName}:/children`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const data = await res.json();
  if (data.error) {
    console.warn(`  [WARN] Could not list "${folderName}": ${data.error.message}`);
    return [];
  }
  return (data.value || []).filter(item => item.name?.toLowerCase().endsWith(".pdf") && item.file);
}

async function downloadFile(token, driveId, itemId) {
  const metaRes  = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const metaData = await metaRes.json();
  if (metaData.error) throw new Error(`Download metadata failed: ${metaData.error.message}`);

  const fileRes = await fetch(metaData["@microsoft.graph.downloadUrl"]);
  if (!fileRes.ok) throw new Error(`Download failed: HTTP ${fileRes.status}`);

  return Buffer.from(await fileRes.arrayBuffer());
}

// ─────────────────────────────────────────────
// PII REDACTION
// ─────────────────────────────────────────────

const PII_PATTERNS = [
  { name: "Email",       pattern: /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g },
  { name: "Phone",       pattern: /(\+?(\d[\s\-.]?){10,14}\d)/g },
  { name: "Aadhaar",    pattern: /\b\d{4}\s?\d{4}\s?\d{4}\b/g },
  { name: "PAN",         pattern: /\b[A-Z]{5}[0-9]{4}[A-Z]{1}\b/g },
  { name: "Passport",    pattern: /\b[A-Z]{1,2}[0-9]{6,7}\b/g },
  { name: "CreditCard",  pattern: /\b(?:\d[ \-]?){15,16}\b/g },
  { name: "SSN",         pattern: /\b\d{3}[- ]?\d{2}[- ]?\d{4}\b/g },
  { name: "DOB",         pattern: /\b(DOB|Date of Birth|D\.O\.B)[:\s]+\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/gi },
  { name: "BankAccount", pattern: /\b\d{9,18}\b/g },
  { name: "IFSC",        pattern: /\b[A-Z]{4}0[A-Z0-9]{6}\b/g },
];

function redactPII(text) {
  let redacted = text;
  const found  = [];
  for (const { name, pattern } of PII_PATTERNS) {
    const matches = redacted.match(pattern);
    if (matches?.length > 0) {
      found.push(`${name} (${matches.length})`);
      redacted = redacted.replace(pattern, "[REDACTED]");
    }
  }
  return { redacted, found };
}

// ─────────────────────────────────────────────
// STEP 1 — Clear entire index
// ─────────────────────────────────────────────

async function clearIndex(searchClient) {
  console.log("\n[CLEAR] Fetching all existing documents from index...");

  const existingIds = [];
  const results = await searchClient.search("*", { select: ["id"], top: 1000 });

  for await (const result of results.results) {
    existingIds.push(result.document.id);
  }

  if (existingIds.length === 0) {
    console.log("[CLEAR] Index is already empty.");
    return;
  }

  console.log(`[CLEAR] Deleting ${existingIds.length} existing document(s)...`);
  const deleteResult = await searchClient.deleteDocuments("id", existingIds);
  const failed = deleteResult.results.filter(r => !r.succeeded);

  if (failed.length > 0) {
    console.warn(`[CLEAR] ${failed.length} document(s) failed to delete.`);
  } else {
    console.log(`[CLEAR] ✓ All ${existingIds.length} old document(s) removed.`);
  }
}

// ─────────────────────────────────────────────
// STEP 2 — Re-index from SharePoint
// ─────────────────────────────────────────────

async function reindex(searchClient, token, driveId) {
  const folders = Object.keys(FOLDER_GROUP_MAP);
  let totalIndexed = 0;
  let totalSkipped = 0;

  for (const folder of folders) {
    const groupIds = FOLDER_GROUP_MAP[folder];
    console.log(`\n─────────────────────────────────`);
    console.log(`[FOLDER] ${folder}/ | Groups: ${groupIds.length}`);

    const files = await listFilesInFolder(token, driveId, folder);

    if (files.length === 0) {
      console.log(`  No PDFs found in this folder.`);
      continue;
    }

    console.log(`  ${files.length} PDF(s) found`);

    for (const file of files) {
      console.log(`\n  [FILE] ${file.name}`);

      try {
        const buffer = await downloadFile(token, driveId, file.id);
        console.log(`  [DOWNLOAD] ${buffer.length} bytes`);

        const pdfData = await pdf(buffer);
        const rawText = pdfData.text;

        if (!rawText || rawText.trim().length === 0) {
          console.log(`  [SKIP] No text extracted — PDF may be image-based`);
          totalSkipped++;
          continue;
        }
        console.log(`  [EXTRACT] ${rawText.length} characters`);

        const { redacted, found } = redactPII(rawText);
        if (found.length > 0) {
          console.log(`  [PII REDACTED] ${found.join(" | ")}`);
        }

        const docId   = Buffer.from(`${folder}__${file.name}`).toString("base64").replace(/[+/=]/g, "_");
        const document = { id: docId, title: file.name, content: redacted, folder, GroupIds: groupIds };

        const result = await searchClient.uploadDocuments([document]);
        const status = result.results[0];

        if (status.succeeded) {
          console.log(`  [INDEXED] ✓ "${file.name}"`);
          totalIndexed++;
        } else {
          console.error(`  [FAILED] "${file.name}" — ${status.errorMessage}`);
          totalSkipped++;
        }
      } catch (err) {
        console.error(`  [ERROR] ${file.name}: ${err.message}`);
        totalSkipped++;
      }
    }
  }

  return { totalIndexed, totalSkipped };
}

// ─────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────

async function sync() {
  if (!SEARCH_SERVICE_NAME || !SEARCH_API_KEY || !TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    console.error("[ERROR] Missing required env vars in local.settings.json");
    process.exit(1);
  }

  const endpoint     = `https://${SEARCH_SERVICE_NAME}.search.windows.net`;
  const credential   = new AzureKeyCredential(SEARCH_API_KEY);
  const searchClient = new SearchClient(endpoint, INDEX_NAME, credential);

  // Step 1 — Wipe the index clean
  await clearIndex(searchClient);

  // Step 2 — Auth + get SharePoint drive
  console.log("\n[AUTH] Getting Microsoft Graph access token...");
  const token = await getAccessToken();
  console.log("[AUTH] Token acquired.");

  const { driveId } = await getSiteAndDrive(token);

  // Step 3 — Re-index all current SharePoint files
  console.log("\n[REINDEX] Starting fresh index from SharePoint...");
  const { totalIndexed, totalSkipped } = await reindex(searchClient, token, driveId);

  console.log(`\n─────────────────────────────────`);
  console.log(`[DONE] Indexed: ${totalIndexed} | Skipped/Failed: ${totalSkipped}`);
  console.log(`Index: "${INDEX_NAME}" on ${SEARCH_SERVICE_NAME}.search.windows.net`);
}

sync();
