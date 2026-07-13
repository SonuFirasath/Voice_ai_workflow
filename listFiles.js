// listFiles.js
const fs = require("fs");

// 1. Load Credentials
const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;
const generalFolder = localSettings.Values.SP_GENERAL_FOLDER;

async function listAccessibleFiles() {
  console.log("1. Authenticating with Microsoft Entra ID...");
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

  if (!tokenResponse.ok) {
    console.error("Auth Failed!");
    return;
  }
  const tokenData = await tokenResponse.json();
  const accessToken = tokenData.access_token;
  console.log("Token secured.\n");

  console.log("2. Scanning the General Policies Folder Index...");

  // We use IsDocument:true to list everything that is a file (ignoring sub-folders)
  const searchQuery = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: { queryString: `IsDocument:true AND path:"${generalFolder}"` },
        // Requesting specific fields makes the output much cleaner to read
        fields: ["name", "webUrl", "lastModifiedDateTime", "createdBy"],
      },
    ],
  };

  const searchResponse = await fetch(
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

  const searchData = await searchResponse.json();
  const hits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];

  console.log("================ INDEXED FILES FOUND ================");
  if (hits.length === 0) {
    console.log(
      "0 files found. (The crawler is still indexing, or the folder is empty).",
    );
  } else {
    hits.forEach((hit, index) => {
      const file = hit.resource;
      console.log(`[${index + 1}] File Name: ${file.name}`);
      console.log(
        `    Last Modified: ${new Date(file.lastModifiedDateTime).toLocaleString()}`,
      );
      console.log(`    Link: ${file.webUrl}`);
      console.log("---------------------------------------------------");
    });
    console.log(`Total Indexed Documents: ${hits.length}`);
  }
}

listAccessibleFiles();
