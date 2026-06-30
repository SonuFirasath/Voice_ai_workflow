// crackTheIndex.js
const fs = require('fs');

const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;

async function crackTheIndex() {
    console.log("1. Authenticating...");
    const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            grant_type: 'client_credentials',
            client_id: clientId,
            client_secret: clientSecret,
            scope: 'https://graph.microsoft.com/.default'
        })
    });
    const accessToken = (await tokenResponse.json()).access_token;
    console.log("Token secured.\n");

    // The GUID extracted from your previous directAccess.js log
    const siteGuid = "939496b9-b57f-440c-90c6-49115a93ed58";
    const exactFolderUrl = "https://chimeratechpvtltd.sharepoint.com/sites/Corporate-Policies/Shared Documents/General_Policies";

    const queriesToTest = [
        // TEST 1: The Site GUID Bypass (Ignores paths completely, searches the whole site)
        `handbook AND SiteId:"${siteGuid}"`,
        
        // TEST 2: The Broad Path Bypass (Searches the whole site via URL wildcard)
        `handbook AND path:"https://chimeratechpvtltd.sharepoint.com/sites/Corporate-Policies*"`,
        
        // TEST 3: The ParentLink Bypass (An alternative to 'path' that sometimes handles spaces better)
        `handbook AND ParentLink:"${exactFolderUrl}"`
    ];

    for (let i = 0; i < queriesToTest.length; i++) {
        console.log(`Executing TEST ${i + 1}...`);
        console.log(`Query: [ ${queriesToTest[i]} ]`);
        
        try {
            const res = await fetch('https://graph.microsoft.com/v1.0/search/query', {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    requests: [{
                        entityTypes: ["driveItem"],
                        query: { queryString: queriesToTest[i] }
                    }]
                })
            });

            const data = await res.json();
            const hits = data.value?.[0]?.hitsContainers?.[0]?.hits || [];

            if (hits.length > 0) {
                console.log(`✅ SUCCESS! Found File: ${hits[0].resource.name}`);
                console.log(`   Snippet: ${hits[0].summary}\n`);
            } else {
                console.log("❌ Failed (0 Results).\n");
            }
        } catch (err) {
            console.log(`❌ Error: ${err.message}\n`);
        }
    }
}

crackTheIndex();