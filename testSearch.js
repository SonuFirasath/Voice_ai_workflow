// testSearch.js
const fs = require("fs");

const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;
const generalFolder = localSettings.Values.SP_GENERAL_FOLDER;

async function runDiagnosticSearch() {
  console.log("1. Authenticating with Microsoft...");
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
  const accessToken = tokenData.access_token;
  console.log("Token secured.\n");

  // TEST 1: Wide open search (ignores folders to check if files are indexed at all)
  console.log("2. Running TEST 1: Broad Search (Is the file indexed yet?)...");
  const test1Query = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: { queryString: "Our Brand Promise" }, // Changed from "leave"
      },
    ],
  };

  const res1 = await fetch("https://graph.microsoft.com/v1.0/search/query", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(test1Query),
  });

  const data1 = await res1.json();
  const hits1 = data1.value?.[0]?.hitsContainers?.[0]?.hits || [];
  console.log(`Results found without path restrictions: ${hits1.length}`);
  if (hits1.length > 0) {
    console.log(`Matched File: ${hits1[0].resource.name}`);
  }

  console.log("\n--------------------------------------------------\n");

  // TEST 2: Folder-restricted search
  console.log("3. Running TEST 2: Folder-Restricted Search...");
  const test2Query = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: {
          queryString: `Our Brand Promise AND (path:"${generalFolder}")`,
        }, // Changed from "leave"
      },
    ],
  };

  const res2 = await fetch("https://graph.microsoft.com/v1.0/search/query", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(test2Query),
  });

  const data2 = await res2.json();
  const hits2 = data2.value?.[0]?.hitsContainers?.[0]?.hits || [];
  console.log(`Results found WITH path restrictions: ${hits2.length}`);
}

runDiagnosticSearch();
