// checkDriveStructure.js
const fs = require("fs");

const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;

async function run() {
  const tokenRes = await fetch(
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
  const token = (await tokenRes.json()).access_token;
  const siteId =
    "chimeratechpvtltd.sharepoint.com,939496b9-b57f-440c-90c6-49115a93ed58,b535fe0e-a060-410d-8e0c-744aaa21e8be";

  // 1. List ALL drives (document libraries) on the site
  console.log("=== ALL DRIVES ON THE SITE ===");
  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    {
      headers: { Authorization: `Bearer ${token}` },
    },
  );
  const drivesData = await drivesRes.json();
  drivesData.value.forEach((d) => {
    console.log(`Drive: "${d.name}" | ID: ${d.id} | webUrl: ${d.webUrl}`);
  });

  // 2. List root children of the default drive
  console.log("\n=== ROOT CHILDREN OF DEFAULT DRIVE ===");
  const rootRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children`,
    {
      headers: { Authorization: `Bearer ${token}` },
    },
  );
  const rootData = await rootRes.json();
  rootData.value.forEach((item) => {
    const icon = item.folder ? "📁" : "📄";
    console.log(`${icon} "${item.name}" | webUrl: ${item.webUrl}`);
    if (item.folder) {
      console.log(`   Children count: ${item.folder.childCount}`);
    }
  });

  // 3. Try alternate path formats to find the folders
  console.log("\n=== TRYING ALTERNATE PATH FORMATS ===");
  const pathAttempts = [
    "General_Policies",
    "Shared Documents/General_Policies",
    "Shared%20Documents/General_Policies",
  ];

  for (const path of pathAttempts) {
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${path}:/children`;
    console.log(`\nTrying path: "/${path}"`);
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.ok) {
      const data = await res.json();
      console.log(`   ✅ SUCCESS! Found ${data.value.length} items:`);
      data.value.forEach((item) => {
        console.log(
          `      ${item.folder ? "📁" : "📄"} ${item.name} (${(item.size / 1024).toFixed(1)} KB)`,
        );
      });
    } else {
      const errText = await res.text();
      console.log(`   ❌ ${res.status}: ${errText.substring(0, 150)}`);
    }
  }

  // 4. Search API with region parameter
  console.log("\n=== SEARCH API WITH REGION PARAMETER ===");
  const regions = ["NAM", "EUR", "APC", "IND"];
  for (const region of regions) {
    const searchBody = {
      requests: [
        {
          entityTypes: ["driveItem"],
          query: { queryString: "IsDocument:true" },
          region: region,
        },
      ],
    };
    const searchRes = await fetch(
      "https://graph.microsoft.com/v1.0/search/query",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(searchBody),
      },
    );
    const searchData = await searchRes.json();
    if (searchData.error) {
      console.log(
        `Region "${region}": ❌ ${searchData.error.message.substring(0, 100)}`,
      );
    } else {
      const hits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];
      const total = searchData.value?.[0]?.hitsContainers?.[0]?.total || 0;
      console.log(
        `Region "${region}": ✅ ${hits.length} hits (total: ${total})`,
      );
      if (hits.length > 0) {
        hits.slice(0, 3).forEach((hit) => {
          console.log(`   📄 ${hit.resource?.name}`);
        });
      }
    }
  }
}

run().catch((err) => console.error("Fatal:", err));
