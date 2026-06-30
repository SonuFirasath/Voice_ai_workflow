// directAccess.js
const fs = require('fs');

const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;

// The exact tenant hostname gathered from previous logs
const tenantHostname = "chimeratechpvtltd.sharepoint.com";
const sitePath = "/sites/Corporate-Policies";

async function testDirectDriveAccess() {
    console.log("1. Authenticating with Microsoft Entra ID...");
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

    const tokenData = await tokenResponse.json();
    const accessToken = tokenData.access_token;
    console.log("Token secured.\n");

    console.log("2. Attempting direct access to the Corporate-Policies site...");
    
    // Step A: Get the Site ID directly
    const siteUrl = `https://graph.microsoft.com/v1.0/sites/${tenantHostname}:${sitePath}`;
    const siteResponse = await fetch(siteUrl, {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    if (!siteResponse.ok) {
        console.error("❌ ACCESS DENIED OR SITE NOT FOUND:");
        console.error(await siteResponse.text());
        console.log("\n-> This means Azure API Permissions (Sites.Read.All) lack Admin Consent, or the site path is wrong.");
        return;
    }

    const siteData = await siteResponse.json();
    console.log(`✅ Site Found! Site ID: ${siteData.id}\n`);

    console.log("3. Fetching root folders inside the Document Library...");
    
    // Step B: Get the default document library (Drive) for this site and list its children
    const driveUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/children`;
    const driveResponse = await fetch(driveUrl, {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });

    if (!driveResponse.ok) {
        console.error("❌ DRIVE ACCESS DENIED:");
        console.error(await driveResponse.text());
        return;
    }

    const driveData = await driveResponse.json();
    const items = driveData.value;

    console.log("================ FOUND FOLDERS/FILES ================");
    if (items.length === 0) {
        console.log("The document library is completely empty.");
    } else {
        items.forEach((item, index) => {
            const type = item.folder ? "📁 FOLDER" : "📄 FILE";
            console.log(`[${index + 1}] ${type}: ${item.name}`);
            if (item.folder) {
                // Print the raw URL path so it can be copied perfectly into local.settings.json
                console.log(`    Internal Web URL: ${item.webUrl}`); 
            }
            console.log("---------------------------------------------------");
        });
    }
}

testDirectDriveAccess();