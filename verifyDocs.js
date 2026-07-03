// verifyDocs.js — Temporarily creates a test index WITHOUT permission filtering 
// to confirm documents exist, then checks who has access.

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

const headers = {
    "Content-Type": "application/json",
    "api-key": SEARCH_KEY
};

async function run() {
    console.log("=".repeat(60));
    console.log("  Document & Permission Verification");
    console.log("=".repeat(60));

    // ── Step 1: Create a temporary test index WITHOUT permission filtering ──
    console.log("\n[1] Creating temporary test index (no permission filtering)...");
    
    // Delete if exists
    await fetch(`${BASE_URL}/indexers/test-indexer?api-version=${API_VERSION}`, { method: 'DELETE', headers });
    await fetch(`${BASE_URL}/indexes/test-index?api-version=${API_VERSION}`, { method: 'DELETE', headers });

    const testIndexPayload = {
        name: "test-index",
        fields: [
            { name: "id", type: "Edm.String", key: true, searchable: false },
            { name: "content", type: "Edm.String", searchable: true, retrievable: true, filterable: false, sortable: false, facetable: false },
            { name: "title", type: "Edm.String", searchable: true, retrievable: true, filterable: false, sortable: false, facetable: false },
            { name: "metadata_spo_item_path", type: "Edm.String", searchable: false, retrievable: true, filterable: false },
            // Make UserIds RETRIEVABLE so we can see what's in them
            { name: "UserIds", type: "Collection(Edm.String)", filterable: true, retrievable: true },
            { name: "GroupIds", type: "Collection(Edm.String)", filterable: true, retrievable: true }
        ]
        // NOTE: No permissionFilterOption — security trimming is OFF
    };

    const idxRes = await fetch(`${BASE_URL}/indexes/test-index?api-version=${API_VERSION}`, {
        method: 'PUT', headers, body: JSON.stringify(testIndexPayload)
    });
    if (!idxRes.ok) {
        console.error("   Failed:", await idxRes.text());
        return;
    }
    console.log("   ✅ Test index created.");

    // ── Step 2: Create a test indexer pointing at the same data source ──
    console.log("\n[2] Creating test indexer...");
    const testIndexerPayload = {
        name: "test-indexer",
        dataSourceName: "sharepoint-datasource",
        targetIndexName: "test-index",
        parameters: {
            maxFailedItems: 50,
            maxFailedItemsPerBatch: 50,
            configuration: {
                indexedFileNameExtensions: ".pdf,.docx,.doc,.txt"
            }
        },
        fieldMappings: [
            {
                sourceFieldName: "metadata_spo_site_asset_item_id",
                targetFieldName: "id",
                mappingFunction: { name: "base64Encode" }
            },
            { sourceFieldName: "metadata_spo_item_name", targetFieldName: "title" },
            { sourceFieldName: "content", targetFieldName: "content" },
            { sourceFieldName: "metadata_spo_item_path", targetFieldName: "metadata_spo_item_path" },
            { sourceFieldName: "metadata_user_ids", targetFieldName: "UserIds" },
            { sourceFieldName: "metadata_group_ids", targetFieldName: "GroupIds" }
        ]
    };

    const indxrRes = await fetch(`${BASE_URL}/indexers/test-indexer?api-version=${API_VERSION}`, {
        method: 'PUT', headers, body: JSON.stringify(testIndexerPayload)
    });
    if (!indxrRes.ok) {
        console.error("   Failed:", await indxrRes.text());
        return;
    }
    console.log("   ✅ Test indexer created and started.");

    // ── Step 3: Wait for indexing ──
    console.log("\n[3] Waiting 45 seconds for indexer to process...");
    await new Promise(r => setTimeout(r, 45000));

    // Check status
    const statusRes = await fetch(`${BASE_URL}/indexers/test-indexer/status?api-version=${API_VERSION}`, { headers });
    const status = await statusRes.json();
    if (status.lastResult) {
        console.log(`   Status: ${status.lastResult.status} | Processed: ${status.lastResult.itemsProcessed} | Failed: ${status.lastResult.itemsFailed}`);
    }

    // ── Step 4: Search the test index (no permission filtering!) ──
    console.log("\n[4] Searching test index (no security trimming)...");
    const searchRes = await fetch(`${BASE_URL}/indexes/test-index/docs/search?api-version=${API_VERSION}`, {
        method: 'POST', headers,
        body: JSON.stringify({
            search: "*",
            top: 10,
            select: "title,UserIds,GroupIds",
            count: true
        })
    });
    const searchData = await searchRes.json();
    console.log(`   @odata.count: ${searchData["@odata.count"]}`);

    if (searchData.value && searchData.value.length > 0) {
        console.log(`   ✅ Found ${searchData.value.length} documents!\n`);
        for (const doc of searchData.value) {
            console.log(`   📄 "${doc.title}"`);
            console.log(`      UserIds: ${JSON.stringify(doc.UserIds)}`);
            console.log(`      GroupIds: ${JSON.stringify(doc.GroupIds)}`);
            console.log();
        }
    } else {
        console.log("   ❌ Still 0 documents. The SharePoint indexer itself may not be pulling content.");
        if (searchData.error) console.log(`   Error: ${JSON.stringify(searchData.error)}`);
    }

    // ── Step 5: Get all users to cross-reference with UserIds ──
    console.log("\n[5] Fetching Entra users to cross-reference...");
    const tokenRes = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            grant_type: 'client_credentials', client_id: CLIENT_ID,
            client_secret: CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default'
        })
    });
    const token = (await tokenRes.json()).access_token;

    // If we found UserIds above, look them up
    if (searchData.value && searchData.value.length > 0) {
        const allUserIds = new Set();
        for (const doc of searchData.value) {
            for (const uid of (doc.UserIds || [])) allUserIds.add(uid);
        }

        if (allUserIds.size > 0) {
            console.log(`   Found ${allUserIds.size} unique UserIds in documents. Looking up names...\n`);
            for (const uid of allUserIds) {
                try {
                    const userRes = await fetch(`https://graph.microsoft.com/v1.0/users/${uid}?$select=displayName,mobilePhone,userPrincipalName`, {
                        headers: { 'Authorization': `Bearer ${token}` }
                    });
                    const userData = await userRes.json();
                    if (userData.displayName) {
                        console.log(`   • ${uid} → ${userData.displayName} (${userData.userPrincipalName}) | Phone: ${userData.mobilePhone || '(none)'}`);
                    } else {
                        console.log(`   • ${uid} → Not found in Entra (may be a group or external user)`);
                    }
                } catch (e) {
                    console.log(`   • ${uid} → Lookup error: ${e.message}`);
                }
            }
        } else {
            console.log("   ⚠️ Documents have EMPTY UserIds — the SharePoint ACLs may use Groups only.");
        }
    }

    // ── Cleanup ──
    console.log("\n[6] Cleaning up test resources...");
    await fetch(`${BASE_URL}/indexers/test-indexer?api-version=${API_VERSION}`, { method: 'DELETE', headers });
    await fetch(`${BASE_URL}/indexes/test-index?api-version=${API_VERSION}`, { method: 'DELETE', headers });
    console.log("   ✅ Test index and indexer deleted.");

    console.log("\n" + "=".repeat(60));
}

run().catch(err => console.error("FATAL:", err));
