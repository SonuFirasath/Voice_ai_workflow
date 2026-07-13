// finalSearchFix.js
const fs = require("fs");

const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;
const generalFolder = localSettings.Values.SP_GENERAL_FOLDER;

async function testFinalSearch() {
  console.log("1. Authenticating...");
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
  const accessToken = (await tokenResponse.json()).access_token;

  console.log("2. Running scoped KQL search with wildcard...");

  // THE FIX: Appending /* tells Microsoft to search the contents of the folder
  const searchScope = `${generalFolder}/*`;

  const queryBody = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: {
          queryString: `handbook AND path:"${searchScope}"`,
        },
      },
    ],
  };

  const res = await fetch("https://graph.microsoft.com/v1.0/search/query", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(queryBody),
  });

  const data = await res.json();
  const hits = data.value?.[0]?.hitsContainers?.[0]?.hits || [];

  console.log(`\nRESULTS FOUND: ${hits.length}`);
  if (hits.length > 0) {
    console.log(`✅ MATCHED FILE: ${hits[0].resource.name}`);
    console.log(`✅ SNIPPET FOR AI: ${hits[0].summary}`);
  } else {
    console.log("❌ Still 0.");
  }
}

testFinalSearch();
