// askQuestion.js — Simulates the full VAPI flow locally
// Usage: node askQuestion.js
// Or with a custom question: node askQuestion.js "What is the leave policy?"

const fs = require('fs');
const { PDFParse } = require('pdf-parse');

const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;
const generalFolder = localSettings.Values.SP_GENERAL_FOLDER;
const managerFolder = localSettings.Values.SP_MANAGER_FOLDER;

const question = process.argv[2] || "What is our brand promise?";

async function askSharePoint() {
    console.log(`\n🔍 Question: "${question}"\n`);

    // 1. Authenticate
    console.log("1. Authenticating with Entra ID...");
    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            grant_type: 'client_credentials',
            client_id: clientId,
            client_secret: clientSecret,
            scope: 'https://graph.microsoft.com/.default'
        })
    });
    const token = (await tokenRes.json()).access_token;
    console.log("   ✅ Token secured.\n");

    // 2. Search SharePoint
    console.log("2. Searching SharePoint (role: manager)...");
    const allowedPaths = `path:"${generalFolder}/*" OR path:"${managerFolder}/*"`;
    const fullQuery = `${question} AND (${allowedPaths})`;

    const searchRes = await fetch('https://graph.microsoft.com/v1.0/search/query', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
            requests: [{
                entityTypes: ["driveItem"],
                query: { queryString: fullQuery },
                region: "IND"
            }]
        })
    });

    const searchData = await searchRes.json();

    if (searchData.error) {
        console.log(`   ❌ Search API Error: ${searchData.error.message}`);
        return;
    }

    const hits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];
    console.log(`   Found: ${hits.length} matching document(s)\n`);

    if (hits.length === 0) {
        console.log("   ❌ No results.");
        return;
    }

    // 3. Download and extract PDF text (same as the webhook now does)
    console.log("3. Downloading & extracting PDF text...\n");
    
    const maxFiles = Math.min(hits.length, 3);
    for (let i = 0; i < maxFiles; i++) {
        const hit = hits[i];
        const fileName = hit.resource?.name || "Unknown";
        const resourceId = hit.resource?.id;
        const driveId = hit.resource?.parentReference?.driveId;

        console.log(`📄 [${i + 1}] ${fileName}`);

        if (!driveId || !resourceId) {
            console.log("   ⚠️  Missing driveId/resourceId — showing search snippet only:");
            console.log(`   ${hit.summary || "(no summary)"}\n`);
            continue;
        }

        const downloadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${resourceId}/content`;
        const downloadRes = await fetch(downloadUrl, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!downloadRes.ok) {
            console.log(`   ❌ Download failed: ${downloadRes.status}\n`);
            continue;
        }

        const fileBuffer = Buffer.from(await downloadRes.arrayBuffer());
        console.log(`   Downloaded: ${(fileBuffer.length / 1024).toFixed(1)} KB`);

        if (fileName.toLowerCase().endsWith('.pdf')) {
            const parser = new PDFParse({ data: fileBuffer });
            const pdfData = await parser.getText();
            await parser.destroy();
            console.log(`   Text: ${pdfData.text.length} chars`);
            console.log("=" .repeat(60));
            // Print first 2000 chars of extracted text
            console.log(pdfData.text.substring(0, 2000));
            console.log("=" .repeat(60));
        }
        console.log("");
    }

    console.log("✅ DONE — The above is the FULL text that GPT-4o will now receive to answer your question.");
}

askSharePoint().catch(err => console.error("Fatal:", err));
