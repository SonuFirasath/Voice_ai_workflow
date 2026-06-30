// verifyFix.js — Tests the exact same query the webhook will use, with both fixes applied
const fs = require('fs');

const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;
const generalFolder = localSettings.Values.SP_GENERAL_FOLDER;
const managerFolder = localSettings.Values.SP_MANAGER_FOLDER;
const salesFolder = localSettings.Values.SP_SALES_FOLDER;

async function run() {
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

    // Simulate a manager searching for "leave" — same as the webhook would do
    const searchQuery = "leave";
    const allowedPaths = `path:"${generalFolder}/*" OR path:"${managerFolder}/*"`;
    const fullQuery = `${searchQuery} AND (${allowedPaths})`;

    console.log("=== SIMULATING WEBHOOK SEARCH (WITH region: IND fix) ===");
    console.log(`Query: ${fullQuery}\n`);

    const searchBody = {
        requests: [{
            entityTypes: ["driveItem"],
            query: { queryString: fullQuery },
            region: "IND"
        }]
    };

    const searchRes = await fetch('https://graph.microsoft.com/v1.0/search/query', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(searchBody)
    });

    const searchData = await searchRes.json();

    if (searchData.error) {
        console.log(`❌ Search API Error: ${searchData.error.message}`);
        return;
    }

    const hits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];
    console.log(`Results: ${hits.length} hit(s)\n`);

    if (hits.length > 0) {
        hits.forEach((hit, i) => {
            console.log(`[${i+1}] 📄 ${hit.resource?.name}`);
            console.log(`    Summary: ${hit.summary}\n`);
        });
        console.log("✅ SUCCESS — This is exactly what VAPI's GPT-4o will receive as context.");
    } else {
        console.log("❌ 0 results. Try different search terms.");
    }

    // Also test each folder individually
    console.log("\n=== PER-FOLDER SEARCH (term: 'policy') ===");
    const folders = { General_Policies: generalFolder, Manager_Policies: managerFolder, Sales_Policies: salesFolder };
    for (const [name, url] of Object.entries(folders)) {
        const body = {
            requests: [{
                entityTypes: ["driveItem"],
                query: { queryString: `IsDocument:true AND path:"${url}/*"` },
                region: "IND"
            }]
        };
        const res = await fetch('https://graph.microsoft.com/v1.0/search/query', {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        const data = await res.json();
        const h = data.value?.[0]?.hitsContainers?.[0]?.hits || [];
        console.log(`${name}: ${h.length} indexed file(s)`);
        h.forEach(hit => console.log(`   📄 ${hit.resource?.name}`));
    }
}

run().catch(err => console.error("Fatal:", err));
