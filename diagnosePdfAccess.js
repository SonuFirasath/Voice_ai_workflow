// diagnosePdfAccess.js — Run with: node diagnosePdfAccess.js
const fs = require("fs");

const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;

const folders = {
  General_Policies: localSettings.Values.SP_GENERAL_FOLDER,
  Manager_Policies: localSettings.Values.SP_MANAGER_FOLDER,
  Sales_Policies: localSettings.Values.SP_SALES_FOLDER,
};

const tenantHostname = "chimeratechpvtltd.sharepoint.com";
const sitePath = "/sites/Corporate-Policies";

async function run() {
  // ==================== STEP 1: AUTHENTICATION ====================
  console.log("=".repeat(60));
  console.log("STEP 1: AUTHENTICATING WITH MICROSOFT ENTRA ID");
  console.log("=".repeat(60));

  const tokenResponse = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
      }),
    },
  );

  const tokenData = await tokenResponse.json();
  if (!tokenData.access_token) {
    console.error(
      "❌ AUTHENTICATION FAILED:",
      JSON.stringify(tokenData, null, 2),
    );
    return;
  }
  const accessToken = tokenData.access_token;
  console.log("✅ Authentication successful.\n");

  // ==================== STEP 2: SITE & DRIVE ACCESS ====================
  console.log("=".repeat(60));
  console.log("STEP 2: VERIFYING SITE & DRIVE ACCESS");
  console.log("=".repeat(60));

  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${tenantHostname}:${sitePath}`;
  const siteResponse = await fetch(siteUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!siteResponse.ok) {
    console.error("❌ CANNOT ACCESS SITE:", await siteResponse.text());
    console.error(
      "   -> Check that Sites.Read.All permission has Admin Consent.",
    );
    return;
  }

  const siteData = await siteResponse.json();
  const siteId = siteData.id;
  console.log(`✅ Site accessible. Site ID: ${siteId}\n`);

  // ==================== STEP 3: LIST FILES IN EACH FOLDER (DIRECT API) ====================
  console.log("=".repeat(60));
  console.log("STEP 3: LISTING FILES IN EACH FOLDER (Direct Graph API)");
  console.log("=".repeat(60));
  console.log(
    "This checks if the service principal can directly browse each folder.\n",
  );

  const folderNames = [
    "General_Policies",
    "Manager_Policies",
    "Sales_Policies",
  ];

  for (const folderName of folderNames) {
    console.log(`--- ${folderName} ---`);
    const listUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/Shared Documents/${folderName}:/children`;

    const listResponse = await fetch(listUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (!listResponse.ok) {
      console.error(
        `   ❌ CANNOT LIST: ${listResponse.status} ${listResponse.statusText}`,
      );
      console.error(`   Response: ${await listResponse.text()}`);
      continue;
    }

    const listData = await listResponse.json();
    const items = listData.value || [];

    if (items.length === 0) {
      console.log("   ⚠️  Folder is EMPTY (no files found).");
    } else {
      let pdfCount = 0;
      items.forEach((item) => {
        const isPdf = item.name?.toLowerCase().endsWith(".pdf");
        if (isPdf) pdfCount++;
        const icon = item.folder ? "📁" : isPdf ? "📕" : "📄";
        const size = item.size ? `(${(item.size / 1024).toFixed(1)} KB)` : "";
        console.log(`   ${icon} ${item.name} ${size}`);
      });
      console.log(`   Total items: ${items.length} | PDF files: ${pdfCount}`);
    }
    console.log("");
  }

  // ==================== STEP 4: SEARCH INDEX TEST (Per Folder) ====================
  console.log("=".repeat(60));
  console.log("STEP 4: SEARCH INDEX TEST (Are PDFs searchable?)");
  console.log("=".repeat(60));
  console.log("This tests if Microsoft's search index has crawled the PDFs.\n");

  for (const [name, folderUrl] of Object.entries(folders)) {
    console.log(`--- ${name} ---`);
    console.log(`   Folder URL: ${folderUrl}`);

    // Test A: List all indexed documents in the folder (with wildcard fix)
    const listQuery = {
      requests: [
        {
          entityTypes: ["driveItem"],
          query: { queryString: `IsDocument:true AND path:"${folderUrl}/*"` },
          fields: ["name", "webUrl", "lastModifiedDateTime"],
        },
      ],
    };

    const listRes = await fetch(
      "https://graph.microsoft.com/v1.0/search/query",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(listQuery),
      },
    );

    const listData = await listRes.json();

    if (listData.error) {
      console.log(`   ❌ SEARCH API ERROR: ${listData.error.message}`);
      continue;
    }

    const hits = listData.value?.[0]?.hitsContainers?.[0]?.hits || [];
    const totalRows = listData.value?.[0]?.hitsContainers?.[0]?.total || 0;
    const moreAvailable =
      listData.value?.[0]?.hitsContainers?.[0]?.moreResultsAvailable || false;

    if (hits.length === 0) {
      console.log(`   ⚠️  0 indexed documents found.`);
      console.log(`   Possible causes:`);
      console.log(
        `     - Files were uploaded recently and haven't been indexed yet (wait 5-15 min).`,
      );
      console.log(
        `     - The folder URL doesn't match what SharePoint uses internally.`,
      );
      console.log(`     - The path filter syntax is wrong.`);

      // Try without path restriction to see if files are indexed anywhere
      const broadQuery = {
        requests: [
          {
            entityTypes: ["driveItem"],
            query: { queryString: `IsDocument:true AND path:"${folderUrl}"` },
          },
        ],
      };
      const broadRes = await fetch(
        "https://graph.microsoft.com/v1.0/search/query",
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(broadQuery),
        },
      );
      const broadData = await broadRes.json();
      const broadHits = broadData.value?.[0]?.hitsContainers?.[0]?.hits || [];
      console.log(
        `   Fallback (without wildcard): ${broadHits.length} result(s)`,
      );
    } else {
      console.log(
        `   ✅ ${hits.length} indexed document(s) found (total: ${totalRows}, more: ${moreAvailable})`,
      );
      hits.forEach((hit) => {
        const name = hit.resource?.name || "Unknown";
        const isPdf = name.toLowerCase().endsWith(".pdf");
        const summaryPreview = (hit.summary || "NO SUMMARY").substring(0, 120);
        console.log(`   ${isPdf ? "📕" : "📄"} ${name}`);
        console.log(`      Summary snippet: "${summaryPreview}..."`);
      });
    }
    console.log("");
  }

  // ==================== STEP 5: CONTENT SEARCH TEST ====================
  console.log("=".repeat(60));
  console.log("STEP 5: CONTENT SEARCH (Can we search INSIDE PDFs?)");
  console.log("=".repeat(60));
  console.log(
    "Searching for a generic term to test full-text search inside documents.\n",
  );

  const testTerms = ["policy", "handbook", "leave", "employee"];

  for (const term of testTerms) {
    const generalFolder = folders.General_Policies;
    const searchQuery = {
      requests: [
        {
          entityTypes: ["driveItem"],
          query: { queryString: `${term} AND path:"${generalFolder}/*"` },
        },
      ],
    };

    const searchRes = await fetch(
      "https://graph.microsoft.com/v1.0/search/query",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(searchQuery),
      },
    );

    const searchData = await searchRes.json();
    const searchHits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];

    if (searchHits.length > 0) {
      console.log(
        `   ✅ "${term}" → ${searchHits.length} hit(s) in General_Policies`,
      );
      console.log(`      Matched: ${searchHits[0].resource?.name}`);
      const snippet = (searchHits[0].summary || "").substring(0, 200);
      console.log(`      Snippet: "${snippet}"`);
    } else {
      console.log(`   ❌ "${term}" → 0 hits in General_Policies`);
    }
  }

  // ==================== STEP 6: PDF DOWNLOAD TEST ====================
  console.log("\n" + "=".repeat(60));
  console.log("STEP 6: PDF DOWNLOAD TEST (Can we download a PDF?)");
  console.log("=".repeat(60));
  console.log(
    "Attempting to download the first PDF found in General_Policies.\n",
  );

  const pdfListUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/Shared Documents/General_Policies:/children`;
  const pdfListRes = await fetch(pdfListUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (pdfListRes.ok) {
    const pdfListData = await pdfListRes.json();
    const firstPdf = (pdfListData.value || []).find((f) =>
      f.name?.toLowerCase().endsWith(".pdf"),
    );

    if (firstPdf) {
      console.log(
        `   Found PDF: ${firstPdf.name} (${(firstPdf.size / 1024).toFixed(1)} KB)`,
      );

      // Download the file content
      const downloadUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${firstPdf.id}/content`;
      const downloadRes = await fetch(downloadUrl, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });

      if (downloadRes.ok) {
        const buffer = Buffer.from(await downloadRes.arrayBuffer());
        console.log(`   ✅ Download successful! Size: ${buffer.length} bytes`);
        console.log(
          `   First 4 bytes (should be %PDF): ${buffer.toString("ascii", 0, 4)}`,
        );

        // Try parsing with pdf-parse if available
        try {
          const pdfParse = require("pdf-parse");
          const pdfData = await pdfParse(buffer);
          console.log(`   ✅ PDF parsed successfully!`);
          console.log(`   Pages: ${pdfData.numpages}`);
          console.log(`   Text length: ${pdfData.text.length} characters`);
          console.log(
            `   First 300 chars: "${pdfData.text.substring(0, 300)}..."`,
          );
        } catch (parseErr) {
          console.log(`   ⚠️  pdf-parse failed: ${parseErr.message}`);
          console.log(
            `   (This means the file downloaded fine, but pdf-parse couldn't read it.)`,
          );
        }
      } else {
        console.error(
          `   ❌ DOWNLOAD FAILED: ${downloadRes.status} ${downloadRes.statusText}`,
        );
        console.error(`   ${await downloadRes.text()}`);
      }
    } else {
      console.log("   ⚠️  No PDF files found in General_Policies folder.");
    }
  }

  // ==================== SUMMARY ====================
  console.log("\n" + "=".repeat(60));
  console.log("DIAGNOSIS COMPLETE");
  console.log("=".repeat(60));
  console.log(`
If STEP 3 shows PDFs but STEP 4 shows 0 indexed documents:
  → SharePoint's search crawler hasn't indexed them yet (wait 5-15 min after upload).
  → Or the folder URL in local.settings.json doesn't match SharePoint's internal path.

If STEP 4 shows documents but STEP 5 shows 0 content hits:
  → The PDFs might be scanned images without OCR (no extractable text).
  → SharePoint can't search inside image-only PDFs.

If STEP 5 works but VAPI still doesn't answer:
  → The snippets returned by search are too short for GPT-4o to give useful answers.
  → Consider downloading the full PDF and extracting text (STEP 6 tests this).
  → Your code currently only sends search 'summary' snippets to the AI, not full document text.
`);
}

run().catch((err) => console.error("Fatal error:", err));
