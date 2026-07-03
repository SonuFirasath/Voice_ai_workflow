// setupSearch.js

const fs = require('fs');
const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
Object.assign(process.env, localSettings.Values);

const SEARCH_SERVICE_NAME = process.env.SEARCH_SERVICE_NAME;
const SEARCH_API_KEY      = process.env.SEARCH_API_KEY;
const SP_SITE_URL         = process.env.SP_SITE_URL;
const ENTRA_TENANT_ID     = process.env.TENANT_ID;
const ENTRA_CLIENT_ID     = process.env.CLIENT_ID;
const ENTRA_CLIENT_SECRET = process.env.CLIENT_SECRET;

const API_VERSION = "2026-05-01-preview";
const BASE_URL    = `https://${SEARCH_SERVICE_NAME}.search.windows.net`;

const headers = {
    "Content-Type": "application/json",
    "api-key": SEARCH_API_KEY
};

async function provisionAzureSearch() {
    try {
        console.log("Starting Azure AI Search Provisioning...");

        // ==========================================
        // STEP 1: Create or Update Data Source
        // ==========================================
        console.log("1. Creating/Updating SharePoint Data Source...");
        const dataSourcePayload = {
            name: "sharepoint-datasource",
            type: "sharepoint",
            indexerPermissionOptions: ["userIds", "groupIds"], 
            credentials: {
                connectionString: `SharePointOnlineEndpoint=${SP_SITE_URL};ApplicationId=${ENTRA_CLIENT_ID};ApplicationSecret=${ENTRA_CLIENT_SECRET};TenantId=${ENTRA_TENANT_ID}`
            },
            container: {
                name: "defaultSiteLibrary"
            }
        };

        const dsRes = await fetch(`${BASE_URL}/datasources/sharepoint-datasource?api-version=${API_VERSION}`, {
            method: 'PUT',
            headers,
            body: JSON.stringify(dataSourcePayload)
        });
        if (!dsRes.ok) throw new Error(`Data Source Error: ${await dsRes.text()}`);
        console.log("   -> Data Source updated successfully.");

        // ==========================================
        // STEP 2: Create or Update Index Schema
        // ==========================================
        console.log("2. Creating/Updating Search Index...");
        const indexPayload = {
            name: "sharepoint-index",
            fields: [
                { name: "id", type: "Edm.String", key: true, searchable: false },
                // CRITICAL: content must NOT be filterable/sortable/facetable — those options
                // force the entire text to be indexed as a single term, which hits the
                // 32,766 byte UTF-8 limit and causes documents to silently fail.
                { name: "content", type: "Edm.String", searchable: true, retrievable: true, filterable: false, sortable: false, facetable: false },
                { name: "title", type: "Edm.String", searchable: true, retrievable: true, filterable: false, sortable: false, facetable: false },
                
                { name: "metadata_spo_item_path", type: "Edm.String", searchable: false, retrievable: true, filterable: false },

                { name: "UserIds", type: "Collection(Edm.String)", filterable: true, retrievable: false, permissionFilter: "userIds" },
                { name: "GroupIds", type: "Collection(Edm.String)", filterable: true, retrievable: false, permissionFilter: "groupIds" }
            ],
            permissionFilterOption: "disabled"
        };

        // Because we are adding a new field, PUT will safely update the schema without breaking it
        const idxRes = await fetch(`${BASE_URL}/indexes/sharepoint-index?api-version=${API_VERSION}`, {
            method: 'PUT',
            headers,
            body: JSON.stringify(indexPayload)
        });
        if (!idxRes.ok) throw new Error(`Index Error: ${await idxRes.text()}`);
        console.log("   -> Index updated successfully.");

        // ==========================================
        // STEP 3: Create or Update Indexer
        // ==========================================
        console.log("3. Creating and Starting Indexer...");
        const indexerPayload = {
            name: "sharepoint-indexer",
            dataSourceName: "sharepoint-datasource",
            targetIndexName: "sharepoint-index",
            parameters: {
                maxFailedItems: 50,         // Ignore up to 50 broken files
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
                { sourceFieldName: "content",                targetFieldName: "content" },
                { sourceFieldName: "metadata_spo_item_path", targetFieldName: "metadata_spo_item_path" },
                { sourceFieldName: "metadata_user_ids",      targetFieldName: "UserIds" },
                { sourceFieldName: "metadata_group_ids",     targetFieldName: "GroupIds" }
            ]
        };

        const indxrRes = await fetch(`${BASE_URL}/indexers/sharepoint-indexer?api-version=${API_VERSION}`, {
            method: 'PUT',
            headers,
            body: JSON.stringify(indexerPayload)
        });
        if (!indxrRes.ok) throw new Error(`Indexer Error: ${await indxrRes.text()}`);
        console.log("   -> Indexer updated successfully.");

        // ==========================================
        // STEP 4: Trigger an immediate indexer run and poll until done
        // ==========================================
        // ==========================================
        // STEP 4 (optional): Reset indexer state for full re-crawl
        // ==========================================
        const forceReset = process.argv.includes('--reset');
        if (forceReset) {
            console.log("4a. Resetting indexer state (--reset flag detected) — forces full re-crawl...");
            const resetRes = await fetch(`${BASE_URL}/indexers/sharepoint-indexer/reset?api-version=${API_VERSION}`, {
                method: 'POST',
                headers
            });
            if (!resetRes.ok && resetRes.status !== 204) {
                console.warn(`   WARNING: Reset returned ${resetRes.status}: ${await resetRes.text()}`);
            } else {
                console.log("   -> Indexer state reset. All documents will be re-crawled and UserIds/GroupIds repopulated.");
            }
        }

        console.log("4. Triggering indexer run now...");
        const runRes = await fetch(`${BASE_URL}/indexers/sharepoint-indexer/run?api-version=${API_VERSION}`, {
            method: 'POST',
            headers
        });
        // 202 Accepted = triggered successfully; 409 = already running (also fine)
        if (runRes.status !== 202 && runRes.status !== 409) {
            throw new Error(`Run trigger failed (${runRes.status}): ${await runRes.text()}`);
        }
        console.log("   -> Indexer run triggered. Polling for completion...");

        // Poll every 5 seconds until the indexer finishes
        let done = false;
        for (let attempt = 1; attempt <= 36; attempt++) {   // max ~3 min
            await new Promise(r => setTimeout(r, 5000));

            const statusRes = await fetch(`${BASE_URL}/indexers/sharepoint-indexer/status?api-version=${API_VERSION}`, { headers });
            const statusData = await statusRes.json();
            const last = statusData.lastResult;

            if (!last) { process.stdout.write('.'); continue; }

            const running = last.status === 'inProgress' || last.endTime === null;
            if (running) {
                process.stdout.write(`\r   -> Still running... (${attempt * 5}s elapsed)`);
                continue;
            }

            console.log(`\n\n   === INDEXER FINISHED ===`);
            console.log(`   Status:         ${last.status}`);
            console.log(`   Docs indexed:   ${last.itemsProcessed}`);
            console.log(`   Docs failed:    ${last.itemsFailed}`);
            if (last.errors && last.errors.length > 0) {
                console.log(`   Errors:`);
                last.errors.slice(0, 5).forEach(e => console.log(`     • ${e.errorMessage || e.message}`));
            }
            done = true;
            break;
        }

        if (!done) {
            console.log("\n   Indexer still running after 3 minutes — check Azure Portal for status.");
        }

        console.log("\nProvisioning complete. Run diagnoseSearch.js to verify documents are indexed.");

    } catch (error) {
        console.error("DEPLOYMENT FAILED:", error.message);
    }
}

provisionAzureSearch();