// checkContent.js — checks if the indexed documents have actual text content
// Run with: node checkContent.js

const fs = require("fs");
const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");

const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
Object.assign(process.env, localSettings.Values);

const endpoint = `https://${process.env.SEARCH_SERVICE_NAME}.search.windows.net`;
const client = new SearchClient(
  endpoint,
  "sharepoint-index",
  new AzureKeyCredential(process.env.SEARCH_API_KEY),
);

async function check() {
  console.log("Fetching all indexed documents and checking content...\n");

  const results = await client.search("*", {
    select: ["title", "content", "metadata_spo_item_path"],
    top: 20,
    includeTotalCount: true,
  });

  let total = 0;
  let emptyCount = 0;

  for await (const r of results.results) {
    total++;
    const doc = r.document;
    const content = doc.content || "";
    const isEmpty = content.trim().length < 20;

    if (isEmpty) emptyCount++;

    console.log(`[${total}] "${doc.title || "(no title)"}"`);
    console.log(`    Path:    ${doc.metadata_spo_item_path || "(none)"}`);
    console.log(
      `    Content: ${
        isEmpty
          ? "❌ EMPTY — no text extracted"
          : `✅ ${content.length} chars — "${content.substring(0, 120).replace(/\s+/g, " ").trim()}..."`
      }`,
    );
    console.log();
  }

  console.log("─".repeat(60));
  console.log(`Total docs: ${total}`);
  console.log(`Empty content: ${emptyCount} / ${total}`);

  if (total === 0) {
    console.log("\n⚠️  No documents returned by wildcard search.");
    console.log(
      "   Even though Azure Portal shows 6 docs, the API returns nothing.",
    );
    console.log(
      "   This is likely caused by permissionFilterOption blocking unfiltered queries.",
    );
    console.log(
      "   → Re-run setupSearch.js after removing permissionFilterOption from the index.",
    );
  } else if (emptyCount === total) {
    console.log("\n❌ ALL DOCUMENTS HAVE EMPTY CONTENT.");
    console.log(
      "   Your PDFs are scanned images with no selectable text layer.",
    );
    console.log(
      "   → Fix: Add OCR cognitive skill to the indexer (tell Claude to do this).",
    );
  } else if (emptyCount > 0) {
    console.log(
      `\n⚠️  ${emptyCount} document(s) have empty content — likely scanned PDFs.`,
    );
  } else {
    console.log("\n✅ All documents have content. Testing keyword search...");

    const testResults = await client.search("policy", {
      select: ["title"],
      top: 5,
    });
    let found = 0;
    for await (const r of testResults.results) {
      found++;
      console.log(`   Match [${found}]: "${r.document.title}"`);
    }
    if (found === 0) {
      console.log(
        "   ⚠️  Keyword 'policy' matched nothing despite content existing.",
      );
      console.log(
        "   → The content field may contain non-searchable/corrupted text.",
      );
    } else {
      console.log(
        `\n✅ Keyword search works — ${found} result(s) for 'policy'.`,
      );
      console.log("   The index and search are working correctly.");
      console.log(
        "   → The problem is in the VAPI/search.js call path, not the index.",
      );
    }
  }
}

check().catch((e) => console.error("Failed:", e.message));
