// src/lib/search.js

const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");
const { getUserGroupIds } = require("./identity");

async function searchSharePoint(searchQuery, callerId, accessToken) {
    console.log(`[SEARCH] query="${searchQuery}" | callerId="${callerId}"`);

    if (!callerId) {
        console.error("[SEARCH] BLOCKED: callerId is null/undefined — caller identity was not resolved.");
        return { error: "Caller identity could not be verified. The phone number may not be registered in the directory." };
    }

    try {
        const serviceName = process.env.SEARCH_SERVICE_NAME;
        const apiKey      = process.env.SEARCH_API_KEY;

        if (!serviceName || !apiKey) {
            console.error("[SEARCH] Missing env vars: SEARCH_SERVICE_NAME or SEARCH_API_KEY");
            return { error: "Search service is not configured." };
        }

        // Get the caller's Entra group memberships
        const groupIds = await getUserGroupIds(accessToken, callerId);

        if (groupIds.length === 0) {
            console.log("[SEARCH] User has no group memberships — access denied to all documents.");
            return { text: "I'm sorry, I could not find any policy documents that you have access to." };
        }

        // Build OData filter: document must be accessible to at least one of the caller's groups
        const groupFilter = groupIds
            .map(id => `GroupIds/any(g: g eq '${id}')`)
            .join(" or ");

        console.log(`[SEARCH] Applying GroupIds filter across ${groupIds.length} group(s)`);

        const endpoint = `https://${serviceName}.search.windows.net`;
        const searchClient = new SearchClient(endpoint, "sharepoint-index", new AzureKeyCredential(apiKey));

        const searchResults = await searchClient.search(searchQuery, {
            filter: groupFilter,
            select: ["title", "content"],
            top: 3
        });

        const extractedTexts = [];

        for await (const result of searchResults.results) {
            const doc     = result.document;
            const title   = doc.title || "Untitled";
            const content = doc.content || "";

            console.log(`[SEARCH] Hit: "${title}" (${content.length} chars)`);

            if (content.length > 0) {
                extractedTexts.push(`[Source: ${title}]\n${content.substring(0, 15000)}`);
            } else {
                extractedTexts.push(`[Source: ${title}]\n(Document found but content is empty)`);
            }
        }

        console.log(`[SEARCH] Returning ${extractedTexts.length} document(s) after group filter.`);

        if (extractedTexts.length === 0) {
            return { text: "No information found in the policy documents you have access to." };
        }

        return { text: extractedTexts.join("\n\n---\n\n") };

    } catch (error) {
        console.error(`[SEARCH] Azure AI Search Error: ${error.message}`);
        console.error(`[SEARCH] Stack: ${error.stack}`);
        return { error: error.message };
    }
}

module.exports = { searchSharePoint };
