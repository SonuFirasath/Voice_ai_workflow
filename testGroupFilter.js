// testGroupFilter.js — verify GroupIds are populated in the Azure AI Search index
// Run with: node testGroupFilter.js

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

const GROUPS = {
  "Policy-General-Readers": "2f996595-97dc-45dd-a85a-0d20759e1682",
  "Policy-Manager-Readers": "70205b81-db45-4fa4-92b7-7fb87a89c73f",
  "Policy-Sales-Readers": "0941d8fd-91c2-415e-8b74-e81cf00d08ea",
};

async function test() {
  console.log("=".repeat(60));
  console.log("  GroupIds Filter Test");
  console.log("=".repeat(60));

  let anyGroupWorked = false;

  for (const [groupName, groupId] of Object.entries(GROUPS)) {
    console.log(`\n[${groupName}]`);
    console.log(`  ID:     ${groupId}`);
    console.log(`  Filter: GroupIds/any(id: id eq '${groupId}')`);

    try {
      const results = await client.search("*", {
        filter: `GroupIds/any(id: id eq '${groupId}')`,
        select: ["title"],
        top: 10,
      });

      let count = 0;
      for await (const r of results.results) {
        count++;
        console.log(`  ✅ "${r.document.title}"`);
      }

      if (count === 0) {
        console.log(
          `  ⚠️  ZERO results — GroupIds not populated for this group`,
        );
      } else {
        anyGroupWorked = true;
      }
    } catch (e) {
      console.log(`  ❌ Error: ${e.message}`);
    }
  }

  console.log("\n" + "=".repeat(60));
  if (anyGroupWorked) {
    console.log(
      "✅ GroupIds ARE populated — ready to implement group-based search filter.",
    );
  } else {
    console.log(
      "❌ GroupIds are NOT populated — the indexer did not extract Entra group ACLs.",
    );
    console.log("\nPossible causes:");
    console.log(
      "  1. The field mapping source name 'metadata_group_ids' may be incorrect",
    );
    console.log(
      "  2. M365 group permissions on SharePoint may not be resolved by the indexer",
    );
    console.log(
      "  3. The app registration may be missing Sites.FullControl.All permission",
    );
  }
  console.log("=".repeat(60));
}

test().catch((err) => console.error("FATAL:", err.message));
