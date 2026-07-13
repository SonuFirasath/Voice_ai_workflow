// debugCallerAccess.js — shows exactly which groups grant a caller access to which documents
// Run with: node debugCallerAccess.js <phone-number>
// Example:  node debugCallerAccess.js +919876543210

const fs = require("fs");
const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");

const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
Object.assign(process.env, localSettings.Values);

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const endpoint = `https://${process.env.SEARCH_SERVICE_NAME}.search.windows.net`;
const client = new SearchClient(
  endpoint,
  "sharepoint-index",
  new AzureKeyCredential(process.env.SEARCH_API_KEY),
);

const phoneArg = process.argv[2];
if (!phoneArg) {
  console.error("Usage: node debugCallerAccess.js <phone-number>");
  process.exit(1);
}

async function getToken() {
  const res = await fetch(
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
  const data = await res.json();
  return data.access_token;
}

async function graphGet(token, url) {
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, ConsistencyLevel: "eventual" },
  });
  return res.json();
}

async function run() {
  console.log(`\nDebugging access for caller: ${phoneArg}\n`);

  const token = await getToken();

  // 1. Find the user by phone number
  const digits = phoneArg.replace(/\D/g, "");
  const variants = [phoneArg, `+${digits}`, digits];
  const conditions = variants
    .map((v) => `mobilePhone eq '${v}' or businessPhones/any(p:p eq '${v}')`)
    .join(" or ");
  const userRes = await graphGet(
    token,
    `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(conditions)}&$select=id,displayName,jobTitle&$count=true`,
  );

  if (!userRes.value || userRes.value.length === 0) {
    console.log("❌ User not found in Entra ID for this phone number.");
    return;
  }

  const user = userRes.value[0];
  console.log(`✅ User found: ${user.displayName} (ID: ${user.id})`);
  console.log(`   Job title: ${user.jobTitle || "(none)"}\n`);

  // 2. Get all group memberships
  const groupRes = await graphGet(
    token,
    `https://graph.microsoft.com/v1.0/users/${user.id}/memberOf?$select=id,displayName&$top=100`,
  );
  const groups = groupRes.value || [];
  console.log(`Group memberships (${groups.length} total):`);
  for (const g of groups) {
    console.log(`  • ${g.displayName || "(unnamed)"} — ${g.id}`);
  }

  // 3. For each group, test which documents it can access in the index
  console.log(`\n${"=".repeat(60)}`);
  console.log("  Access via each group:");
  console.log("=".repeat(60));

  const allAccessibleDocs = new Set();

  for (const g of groups) {
    const filter = `GroupIds/any(x: x eq '${g.id}')`;
    const results = await client.search("*", {
      filter,
      select: ["title"],
      top: 10,
    });

    const docs = [];
    for await (const r of results.results) {
      docs.push(r.document.title);
      allAccessibleDocs.add(r.document.title);
    }

    if (docs.length > 0) {
      console.log(`\n  [${g.displayName || g.id}]`);
      docs.forEach((d) => console.log(`    ✅ ${d}`));
    }
  }

  console.log(`\n${"=".repeat(60)}`);
  console.log("  Summary — all documents this caller can access:");
  console.log("=".repeat(60));
  if (allAccessibleDocs.size === 0) {
    console.log("  ❌ None — caller has no access to any policy documents.");
  } else {
    for (const doc of allAccessibleDocs) {
      console.log(`  ✅ ${doc}`);
    }
  }
  console.log();
}

run().catch((err) => console.error("FATAL:", err.message));
