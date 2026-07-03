// deepDiagnose.js — Tests search bypassing the permission filter to confirm docs exist,
// then tests with specific user IDs to find the right format.

const fs = require('fs');
const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
Object.assign(process.env, localSettings.Values);

const SEARCH_SERVICE = process.env.SEARCH_SERVICE_NAME;
const SEARCH_KEY     = process.env.SEARCH_API_KEY;
const TENANT_ID      = process.env.TENANT_ID;
const CLIENT_ID      = process.env.CLIENT_ID;
const CLIENT_SECRET  = process.env.CLIENT_SECRET;
const API_VERSION    = "2026-05-01-preview";
const BASE_URL       = `https://${SEARCH_SERVICE}.search.windows.net`;

async function run() {
    console.log("=".repeat(60));
    console.log("  Deep Search Diagnostic");
    console.log("=".repeat(60));

    // ── Test 1: Get document count from index stats ──
    console.log("\n[1] Checking index statistics...");
    const statsRes = await fetch(`${BASE_URL}/indexes/sharepoint-index?api-version=${API_VERSION}`, {
        headers: { "api-key": SEARCH_KEY }
    });
    const statsData = await statsRes.json();
    console.log(`   Document count (from schema): Check Azure Portal for exact count.`);

    // ── Test 2: Search via REST API with explicit $count ──
    console.log("\n[2] Searching via REST API (no filter, requesting $count)...");
    const searchRes = await fetch(
        `${BASE_URL}/indexes/sharepoint-index/docs?api-version=${API_VERSION}&search=*&$count=true&$top=5&$select=title`,
        { headers: { "api-key": SEARCH_KEY } }
    );
    const searchData = await searchRes.json();
    console.log(`   @odata.count: ${searchData["@odata.count"]}`);
    console.log(`   Results returned: ${searchData.value?.length || 0}`);
    if (searchData.value) {
        for (const doc of searchData.value) {
            console.log(`   • "${doc.title}"`);
        }
    }
    if (searchData.error) {
        console.log(`   Error: ${JSON.stringify(searchData.error)}`);
    }

    // ── Test 3: Search via POST with no security trimming ──
    console.log("\n[3] Searching via POST (no filter)...");
    const postSearchRes = await fetch(
        `${BASE_URL}/indexes/sharepoint-index/docs/search?api-version=${API_VERSION}`,
        {
            method: 'POST',
            headers: { "api-key": SEARCH_KEY, "Content-Type": "application/json" },
            body: JSON.stringify({
                search: "*",
                top: 5,
                select: "title,metadata_spo_item_path",
                count: true
            })
        }
    );
    const postData = await postSearchRes.json();
    console.log(`   @odata.count: ${postData["@odata.count"]}`);
    console.log(`   Results: ${postData.value?.length || 0}`);
    if (postData.value) {
        for (const doc of postData.value) {
            console.log(`   • "${doc.title}" → ${doc.metadata_spo_item_path || '(no path)'}`);
        }
    }
    if (postData.error) {
        console.log(`   Error: ${JSON.stringify(postData.error)}`);
    }

    // ── Test 4: Get Entra token and find the test user's ID ──
    console.log("\n[4] Getting Entra token to look up all users...");
    const tokenRes = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            grant_type: 'client_credentials',
            client_id: CLIENT_ID,
            client_secret: CLIENT_SECRET,
            scope: 'https://graph.microsoft.com/.default'
        })
    });
    const tokenData = await tokenRes.json();
    const accessToken = tokenData.access_token;

    // Get the site owner / users who have access to the SharePoint site
    console.log("\n[5] Checking SharePoint site permissions (who has access to the Corporate-Policies site)...");
    
    // Try getting site members
    try {
        const siteRes = await fetch(
            `https://graph.microsoft.com/v1.0/sites/chimeratechpvtltd.sharepoint.com:/sites/Corporate-Policies`,
            { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );
        const siteData = await siteRes.json();
        if (siteData.id) {
            console.log(`   Site ID: ${siteData.id}`);
            
            // Get site permissions
            const permsRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteData.id}/permissions`,
                { headers: { 'Authorization': `Bearer ${accessToken}` } }
            );
            const permsData = await permsRes.json();
            if (permsData.value) {
                console.log(`   Site permissions (${permsData.value.length}):`);
                for (const perm of permsData.value) {
                    const identity = perm.grantedToIdentitiesV2 || perm.grantedToIdentities || [];
                    console.log(`   • ${JSON.stringify(perm.roles)} → ${JSON.stringify(identity).substring(0, 200)}`);
                }
            } else {
                console.log(`   Permissions response: ${JSON.stringify(permsData).substring(0, 300)}`);
            }
        } else {
            console.log(`   Site lookup: ${JSON.stringify(siteData).substring(0, 300)}`);
        }
    } catch(e) {
        console.log(`   Error: ${e.message}`);
    }

    // ── Test 6: Try search with the security filter using a known user ──
    console.log("\n[6] Testing search POST with security filter for first user...");
    const testUserId = "0fd5007d-6af1-4bb3-8469-c995a3ba6ee4"; // Aakarsh
    const filteredRes = await fetch(
        `${BASE_URL}/indexes/sharepoint-index/docs/search?api-version=${API_VERSION}`,
        {
            method: 'POST',
            headers: { "api-key": SEARCH_KEY, "Content-Type": "application/json" },
            body: JSON.stringify({
                search: "*",
                filter: `UserIds/any(id: id eq '${testUserId}')`,
                top: 5,
                select: "title",
                count: true
            })
        }
    );
    const filteredData = await filteredRes.json();
    console.log(`   Results with UserIds filter: ${filteredData.value?.length || 0}`);
    if (filteredData.value) {
        for (const doc of filteredData.value) {
            console.log(`   • "${doc.title}"`);
        }
    }
    if (filteredData.error) {
        console.log(`   Error: ${JSON.stringify(filteredData.error)}`);
    }

    // ── Test 7: Try search with GroupIds filter ──
    console.log("\n[7] Testing search with GroupIds filter (using common group patterns)...");
    // Try with the Everyone group (a common default)
    const groupFilterRes = await fetch(
        `${BASE_URL}/indexes/sharepoint-index/docs/search?api-version=${API_VERSION}`,
        {
            method: 'POST',
            headers: { "api-key": SEARCH_KEY, "Content-Type": "application/json" },
            body: JSON.stringify({
                search: "*",
                filter: `GroupIds/any(g: g eq 'c:0(.s|true')`,
                top: 5,
                select: "title",
                count: true
            })
        }
    );
    const groupData = await groupFilterRes.json();
    console.log(`   Results with 'Everyone' group filter: ${groupData.value?.length || 0}`);
    if (groupData.value) {
        for (const doc of groupData.value) {
            console.log(`   • "${doc.title}"`);
        }
    }
    if (groupData.error) {
        console.log(`   Error: ${JSON.stringify(groupData.error)}`);
    }

    // ── Test 8: Check indexer detailed status ──
    console.log("\n[8] Indexer detailed execution history...");
    const histRes = await fetch(
        `${BASE_URL}/indexers/sharepoint-indexer/status?api-version=${API_VERSION}`,
        { headers: { "api-key": SEARCH_KEY } }
    );
    const histData = await histRes.json();
    if (histData.executionHistory) {
        for (const exec of histData.executionHistory.slice(0, 3)) {
            console.log(`   Run: ${exec.status} | Processed: ${exec.itemsProcessed} | Failed: ${exec.itemsFailed} | Start: ${exec.startTime}`);
            if (exec.warnings?.length > 0) {
                for (const w of exec.warnings.slice(0, 3)) {
                    console.log(`     ⚠️ Warning: ${w.message?.substring(0, 200)}`);
                }
            }
        }
    }

    console.log("\n" + "=".repeat(60));
}

run().catch(err => console.error("FATAL:", err));
