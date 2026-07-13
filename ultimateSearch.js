// ultimateSearch.js
const fs = require("fs");

const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;

async function executeUltimateSearch() {
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

  console.log("2. Running the flawless KQL query...");

  // THE FIX: Exact spelling, normal space, and an asterisk at the end inside the quotes
  const exactKQL = `handbook AND path:"https://chimeratechpvtltd.sharepoint.com/sites/Corporate-Policies/Shared Documents/General_Policies*"`;

  console.log(`Executing Query: [ ${exactKQL} ]\n`);

  const queryBody = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: { queryString: exactKQL },
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

  if (hits.length > 0) {
    console.log(`✅ SUCCESS! AI CAN READ THIS:`);
    console.log(`   File: ${hits[0].resource.name}`);
    console.log(`   Snippet: ${hits[0].summary}`);
  } else {
    console.log(
      "❌ Still 0. (We need to check the exact SharePoint list name).",
    );
  }
}

executeUltimateSearch();
