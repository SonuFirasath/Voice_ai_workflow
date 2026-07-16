// src/lib/search.js

// We need the Graph access token to authenticate the search
const { getAccessToken } = require('./auth');

async function searchSharePoint(searchQuery, callerId) {
    // callerId is intentionally ignored for this demo to search tenant-wide
    console.log(`[SEARCH] Demo Mode: Bypassing Azure AI. Querying Graph API for "${searchQuery}"`);

    try {
        const accessToken = await getAccessToken();

        if (!accessToken) {
            console.error("[SEARCH] No access token available.");
            return { error: "Authentication with Microsoft Graph failed." };
        }

        // The global Microsoft Graph Search endpoint
        const endpoint = `https://graph.microsoft.com/v1.0/search/query`;
        
        const requestBody = {
            requests: [
                {
                    // Targets SharePoint document libraries and lists
                    entityTypes: ["driveItem", "listItem"],
                    query: {
                        queryString: searchQuery
                    },
                    // We request the name and URL; Graph returns a 'summary' snippet automatically
                    fields: ["name", "title", "webUrl"]
                }
            ]
        };

        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestBody)
        });

        const data = await response.json();

        if (data.error) {
            console.error(`[SEARCH] Graph API Error:`, data.error);
            return { error: data.error.message };
        }

        // Drill down into the Graph API response structure to find the hits
        const hits = data.value?.[0]?.hitsContainers?.[0]?.hits || [];
        
        if (hits.length === 0) {
            return { text: "No information found in the SharePoint directory." };
        }

        const extractedTexts = [];
        let docCount = 0;

        // Loop through the top 3 hits
        for (const hit of hits.slice(0, 3)) { 
            docCount++;
            const resource = hit.resource;
            const title = resource.name || resource.title || "Untitled";
            
            // Graph Search returns an HTML-formatted snippet highlighting the search match
            const summary = hit.summary || "(No text summary available)";
            
            // Microsoft inserts <c0> tags to highlight words. This cleans them out.
            const cleanSummary = summary.replace(/<[^>]*>?/gm, '');

            console.log(`[SEARCH] Hit ${docCount}: "${title}"`);
            
            extractedTexts.push(`[Source: ${title}]\nSummary: ${cleanSummary}`);
        }

        return { text: extractedTexts.join("\n\n---\n\n") };

    } catch (error) {
        console.error(`[SEARCH] Microsoft Graph API Error: ${error.message}`);
        console.error(`[SEARCH] Stack: ${error.stack}`);
        return { error: error.message };
    }
}

module.exports = { searchSharePoint };