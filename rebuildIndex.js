// rebuildIndex.js — Deletes and recreates the index + indexer with the fixed schema, then starts re-indexing.
// Run: node rebuildIndex.js

const fs = require("fs");
const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
Object.assign(process.env, localSettings.Values);

const SEARCH_SERVICE = process.env.SEARCH_SERVICE_NAME;
const SEARCH_KEY = process.env.SEARCH_API_KEY;
const API_VERSION = "2026-05-01-preview";
const BASE_URL = `https://${SEARCH_SERVICE}.search.windows.net`;

const headers = {
  "Content-Type": "application/json",
  "api-key": SEARCH_KEY,
};

async function deleteIfExists(resourceType, name) {
  const url = `${BASE_URL}/${resourceType}/${name}?api-version=${API_VERSION}`;
  const res = await fetch(url, { method: "DELETE", headers });
  if (res.ok || res.status === 404) {
    console.log(`   ✅ ${resourceType}/${name} deleted (or didn't exist).`);
  } else {
    const text = await res.text();
    console.error(
      `   ❌ Failed to delete ${resourceType}/${name}: ${res.status} ${text}`,
    );
  }
}

async function rebuild() {
  console.log("=".repeat(60));
  console.log("  Rebuilding Azure AI Search Index (Fixed Schema)");
  console.log("=".repeat(60));

  // ── Step 1: Delete indexer first (depends on index + datasource) ──
  console.log("\n[1] Deleting old indexer...");
  await deleteIfExists("indexers", "sharepoint-indexer");

  // ── Step 2: Delete old index ──
  console.log("\n[2] Deleting old index...");
  await deleteIfExists("indexes", "sharepoint-index");

  // We keep the data source — it doesn't need changing.
  // If you need to recreate it too, just run setupSearch.js after this.

  // ── Step 3: Recreate index with FIXED schema ──
  console.log(
    "\n[3] Creating index with FIXED schema (filterable:false on content)...",
  );
  const indexPayload = {
    name: "sharepoint-index",
    fields: [
      { name: "id", type: "Edm.String", key: true, searchable: false },
      // CRITICAL: These MUST have filterable/sortable/facetable = false
      // Otherwise Azure tries to index the entire text as one term → 32766 byte crash
      {
        name: "content",
        type: "Edm.String",
        searchable: true,
        retrievable: true,
        filterable: false,
        sortable: false,
        facetable: false,
      },
      {
        name: "title",
        type: "Edm.String",
        searchable: true,
        retrievable: true,
        filterable: false,
        sortable: false,
        facetable: false,
      },
      {
        name: "metadata_spo_item_path",
        type: "Edm.String",
        searchable: false,
        retrievable: true,
        filterable: false,
      },
      {
        name: "UserIds",
        type: "Collection(Edm.String)",
        filterable: true,
        retrievable: false,
        permissionFilter: "userIds",
      },
      {
        name: "GroupIds",
        type: "Collection(Edm.String)",
        filterable: true,
        retrievable: false,
        permissionFilter: "groupIds",
      },
    ],
    permissionFilterOption: "enabled",
  };

  const idxRes = await fetch(
    `${BASE_URL}/indexes/sharepoint-index?api-version=${API_VERSION}`,
    {
      method: "PUT",
      headers,
      body: JSON.stringify(indexPayload),
    },
  );
  if (!idxRes.ok) {
    const text = await idxRes.text();
    console.error(`   ❌ Failed to create index: ${text}`);
    return;
  }
  console.log("   ✅ Index created with fixed schema.");

  // ── Step 4: Recreate indexer ──
  console.log("\n[4] Creating indexer...");
  const SP_SITE_URL = process.env.SP_SITE_URL;
  const TENANT_ID = process.env.TENANT_ID;
  const CLIENT_ID = process.env.CLIENT_ID;
  const CLIENT_SECRET = process.env.CLIENT_SECRET;

  const indexerPayload = {
    name: "sharepoint-indexer",
    dataSourceName: "sharepoint-datasource",
    targetIndexName: "sharepoint-index",
    parameters: {
      maxFailedItems: 50,
      maxFailedItemsPerBatch: 50,
      configuration: {
        indexedFileNameExtensions: ".pdf,.docx,.doc,.txt",
      },
    },
    fieldMappings: [
      {
        sourceFieldName: "metadata_spo_site_asset_item_id",
        targetFieldName: "id",
        mappingFunction: { name: "base64Encode" },
      },
      { sourceFieldName: "metadata_spo_item_name", targetFieldName: "title" },
      { sourceFieldName: "content", targetFieldName: "content" },
      {
        sourceFieldName: "metadata_spo_item_path",
        targetFieldName: "metadata_spo_item_path",
      },
      // CRITICAL: Map SharePoint ACL fields to index permission fields
      { sourceFieldName: "metadata_user_ids", targetFieldName: "UserIds" },
      { sourceFieldName: "metadata_group_ids", targetFieldName: "GroupIds" },
    ],
  };

  const indxrRes = await fetch(
    `${BASE_URL}/indexers/sharepoint-indexer?api-version=${API_VERSION}`,
    {
      method: "PUT",
      headers,
      body: JSON.stringify(indexerPayload),
    },
  );
  if (!indxrRes.ok) {
    const text = await indxrRes.text();
    console.error(`   ❌ Failed to create indexer: ${text}`);
    return;
  }
  console.log("   ✅ Indexer created and started.");

  // ── Step 5: Wait and check status ──
  console.log("\n[5] Waiting 30 seconds for indexer to process...");
  await new Promise((r) => setTimeout(r, 30000));

  const statusRes = await fetch(
    `${BASE_URL}/indexers/sharepoint-indexer/status?api-version=${API_VERSION}`,
    {
      headers,
    },
  );
  const status = await statusRes.json();
  if (status.lastResult) {
    const lr = status.lastResult;
    console.log(`   Status: ${lr.status}`);
    console.log(`   Items processed: ${lr.itemsProcessed}`);
    console.log(`   Items failed: ${lr.itemsFailed}`);
    if (lr.errors?.length > 0) {
      console.log(`   Errors:`);
      for (const err of lr.errors.slice(0, 3)) {
        console.log(`     • ${err.errorMessage || err.message}`);
      }
    }
  }

  // ── Step 6: Quick search test ──
  console.log("\n[6] Testing search (no filter)...");
  const {
    SearchClient,
    AzureKeyCredential,
  } = require("@azure/search-documents");
  const client = new SearchClient(
    `${BASE_URL}`,
    "sharepoint-index",
    new AzureKeyCredential(SEARCH_KEY),
  );

  try {
    const results = await client.search("*", { select: ["title"], top: 10 });
    let count = 0;
    for await (const r of results.results) {
      count++;
      console.log(`   • Doc ${count}: "${r.document.title}"`);
    }
    if (count === 0) {
      console.log(
        "   ⚠️ Still empty — indexer may need more time. Try running diagnoseSearch.js in a few minutes.",
      );
    } else {
      console.log(`   ✅ ${count} documents found! Index is working.`);
    }
  } catch (e) {
    console.error("   Search test error:", e.message);
  }

  console.log("\n" + "=".repeat(60));
  console.log(
    "  Done. If index is still empty, wait 2-3 minutes and run diagnoseSearch.js",
  );
  console.log("=".repeat(60));
}

rebuild().catch((err) => console.error("FATAL:", err));
