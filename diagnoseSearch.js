// diagnoseSearch.js — Run this to find out exactly why search returns no results

const fs = require("fs");
const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");

const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
Object.assign(process.env, localSettings.Values);

const SEARCH_SERVICE = process.env.SEARCH_SERVICE_NAME;
const SEARCH_KEY = process.env.SEARCH_API_KEY;
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const endpoint = `https://${SEARCH_SERVICE}.search.windows.net`;
const API_VERSION = "2026-05-01-preview";

async function run() {
  console.log("=".repeat(70));
  console.log("  VAPI + Azure AI Search — Full Diagnostic");
  console.log("=".repeat(70));

  // ──────────────────────────────────────────────
  // STEP 1: Get Entra ID token
  // ──────────────────────────────────────────────
  console.log("\n[1] Authenticating with Entra ID...");
  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default",
      }),
    },
  );
  const tokenData = await tokenRes.json();
  if (!tokenData.access_token) {
    console.error(
      "   FAILED:",
      tokenData.error_description || JSON.stringify(tokenData),
    );
    return;
  }
  const accessToken = tokenData.access_token;
  console.log("   ✅ Token acquired.");

  // ──────────────────────────────────────────────
  // STEP 2: Look up YOUR phone number in Entra ID
  // ──────────────────────────────────────────────
  console.log("\n[2] Looking up users with phone numbers in Entra ID...");
  const usersRes = await fetch(
    `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mobilePhone,businessPhones,jobTitle&$top=20`,
    { headers: { Authorization: `Bearer ${accessToken}` } },
  );
  const usersData = await usersRes.json();
  if (usersData.value) {
    console.log(`   Found ${usersData.value.length} users:`);
    for (const u of usersData.value) {
      const phones = [u.mobilePhone, ...(u.businessPhones || [])].filter(
        Boolean,
      );
      console.log(
        `   • ${u.displayName} | ID: ${u.id} | Phones: ${phones.length > 0 ? phones.join(", ") : "(none)"} | Title: ${u.jobTitle || "(none)"}`,
      );
    }
  } else {
    console.log("   No users found or error:", JSON.stringify(usersData));
  }

  // ──────────────────────────────────────────────
  // STEP 3: Check the index schema
  // ──────────────────────────────────────────────
  console.log("\n[3] Checking index schema...");
  const schemaRes = await fetch(
    `${endpoint}/indexes/sharepoint-index?api-version=${API_VERSION}`,
    {
      headers: { "api-key": SEARCH_KEY },
    },
  );
  const schema = await schemaRes.json();
  if (schema.error) {
    console.error("   ERROR:", schema.error.message);
  } else {
    console.log("   Index fields:");
    for (const f of schema.fields) {
      console.log(
        `   • ${f.name} (${f.type}) | filterable: ${f.filterable || false} | permissionFilter: ${f.permissionFilter || "none"}`,
      );
    }
  }

  // ──────────────────────────────────────────────
  // STEP 4: Search WITHOUT security filter (does the index have data?)
  // ──────────────────────────────────────────────
  console.log(
    "\n[4] Searching WITHOUT security filter (testing if index has data)...",
  );
  const client = new SearchClient(
    endpoint,
    "sharepoint-index",
    new AzureKeyCredential(SEARCH_KEY),
  );

  let totalDocs = 0;
  try {
    const noFilterResults = await client.search("*", {
      select: ["title", "content"],
      top: 10,
    });
    for await (const r of noFilterResults.results) {
      totalDocs++;
      const content = r.document.content || "";
      const preview =
        content.length > 0
          ? content.substring(0, 150).replace(/\s+/g, " ")
          : "(EMPTY — no text extracted)";
      console.log(`   • Doc ${totalDocs}: "${r.document.title}"`);
      console.log(`     Content (${content.length} chars): ${preview}`);
    }
    if (totalDocs === 0) {
      console.log("   ⚠️  INDEX IS EMPTY — No documents found at all!");
      console.log(
        "   → Indexer may not have completed, or SharePoint auth failed.",
      );
      console.log("   → Check step [7] indexer status below for errors.");
    } else {
      const emptyDocs = await (async () => {
        let empty = 0;
        const r2 = await client.search("*", { select: ["content"], top: 50 });
        for await (const r of r2.results) {
          if (!r.document.content || r.document.content.length < 10) empty++;
        }
        return empty;
      })();
      if (emptyDocs === totalDocs) {
        console.log(
          `\n   ⚠️  ALL ${totalDocs} DOCS HAVE EMPTY CONTENT FIELDS!`,
        );
        console.log(
          "   → PDFs are almost certainly scanned images with no text layer.",
        );
        console.log(
          "   → Fix: Add OCR cognitive skill to the indexer (tell Claude to do this).",
        );
      } else {
        console.log(
          `\n   ✅ ${totalDocs} documents indexed, content appears to be extracted.`,
        );
      }
    }
  } catch (e) {
    console.error("   ERROR:", e.message);
  }

  // ──────────────────────────────────────────────
  // STEP 5: Check what's in the UserIds field
  // ──────────────────────────────────────────────
  console.log("\n[5] Checking UserIds field values in indexed documents...");
  try {
    // Try to retrieve UserIds — it might not be retrievable
    const userIdResults = await client.search("*", {
      select: ["title", "UserIds", "GroupIds"],
      top: 10,
    });
    let docCount = 0;
    for await (const r of userIdResults.results) {
      docCount++;
      const doc = r.document;
      const userIds = doc.UserIds || [];
      const groupIds = doc.GroupIds || [];
      console.log(
        `   • "${doc.title}" → UserIds: [${userIds.join(", ")}] | GroupIds: [${groupIds.join(", ")}]`,
      );
    }
    if (docCount === 0) {
      console.log("   No documents returned.");
    }
  } catch (e) {
    console.log(
      `   ⚠️  Could not retrieve UserIds (may be non-retrievable): ${e.message}`,
    );
    console.log("   → Trying via REST API directly...");

    // Try the REST API to check a doc
    try {
      const restRes = await fetch(
        `${endpoint}/indexes/sharepoint-index/docs?api-version=${API_VERSION}&search=*&$top=5&$select=title,UserIds,GroupIds`,
        {
          headers: {
            "api-key": SEARCH_KEY,
            "Content-Type": "application/json",
          },
        },
      );
      const restData = await restRes.json();
      if (restData.value) {
        for (const doc of restData.value) {
          console.log(
            `   • "${doc.title}" → UserIds: ${JSON.stringify(doc.UserIds)} | GroupIds: ${JSON.stringify(doc.GroupIds)}`,
          );
        }
      } else {
        console.log("   REST also failed:", JSON.stringify(restData));
      }
    } catch (e2) {
      console.log("   REST fallback also failed:", e2.message);
    }
  }

  // ──────────────────────────────────────────────
  // STEP 6: Test search WITH a specific user's ID
  // ──────────────────────────────────────────────
  if (usersData.value && usersData.value.length > 0) {
    const testUser = usersData.value[0];
    console.log(
      `\n[6] Testing search WITH security filter for: ${testUser.displayName} (${testUser.id})...`,
    );

    const filter = `UserIds/any(id: id eq '${testUser.id}')`;
    console.log(`   Filter: ${filter}`);

    try {
      const filteredResults = await client.search("*", {
        filter: filter,
        select: ["title"],
        top: 10,
      });
      let fCount = 0;
      for await (const r of filteredResults.results) {
        fCount++;
        console.log(`   • Doc ${fCount}: "${r.document.title}"`);
      }
      if (fCount === 0) {
        console.log("   ⚠️  ZERO results with security filter!");
        console.log(
          "   → This means the UserIds field does NOT contain this user's Entra Object ID.",
        );
        console.log(
          "   → The SharePoint indexer may not have populated ACLs, or the user doesn't have permission to the SharePoint files.",
        );
      } else {
        console.log(
          `   ✅ Found ${fCount} documents accessible to ${testUser.displayName}.`,
        );
      }
    } catch (e) {
      console.error("   Filter search error:", e.message);
    }
  }

  // ──────────────────────────────────────────────
  // STEP 7: Check indexer status
  // ──────────────────────────────────────────────
  console.log("\n[7] Checking indexer status...");
  try {
    const indexerStatusRes = await fetch(
      `${endpoint}/indexers/sharepoint-indexer/status?api-version=${API_VERSION}`,
      { headers: { "api-key": SEARCH_KEY } },
    );
    const indexerStatus = await indexerStatusRes.json();
    if (indexerStatus.lastResult) {
      const lr = indexerStatus.lastResult;
      console.log(`   Status: ${lr.status}`);
      console.log(`   Items processed: ${lr.itemsProcessed}`);
      console.log(`   Items failed: ${lr.itemsFailed}`);
      console.log(`   Start time: ${lr.startTime}`);
      console.log(`   End time: ${lr.endTime}`);
      if (lr.errors && lr.errors.length > 0) {
        console.log(`   Errors:`);
        for (const err of lr.errors.slice(0, 5)) {
          console.log(`     • ${err.message || JSON.stringify(err)}`);
        }
      }
    } else {
      console.log("   No run history found.");
    }
  } catch (e) {
    console.error("   Error checking indexer:", e.message);
  }

  console.log("\n" + "=".repeat(70));
  console.log("  Diagnostic complete. Review the output above.");
  console.log("=".repeat(70));
}

run().catch((err) => console.error("FATAL:", err));
