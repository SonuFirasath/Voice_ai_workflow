// src/lib/search.js

const { SearchClient, AzureKeyCredential } = require("@azure/search-documents");

async function searchSharePoint(searchQuery, callerId) {
    console.log(`[SEARCH] Called with query="${searchQuery}", callerId="${callerId}"`);

    if (!callerId) {
        console.error("[SEARCH] BLOCKED: callerId is null/undefined — caller identity was not resolved.");
        return { error: "Caller identity could not be verified. The phone number may not be registered in the directory." };
    }

    try {
        const serviceName = process.env.SEARCH_SERVICE_NAME;
        const apiKey = process.env.SEARCH_API_KEY;

        if (!serviceName || !apiKey) {
            console.error("[SEARCH] Missing env vars: SEARCH_SERVICE_NAME or SEARCH_API_KEY");
            return { error: "Search service is not configured." };
        }

        const endpoint = `https://${serviceName}.search.windows.net`;
        const searchClient = new SearchClient(
            endpoint,
            "sharepoint-index",
            new AzureKeyCredential(apiKey)
        );

        const securityFilter = `UserIds/any(id: id eq '${callerId}')`;
        console.log(`[SEARCH] Endpoint: ${endpoint}`);
        console.log(`[SEARCH] Security filter: ${securityFilter}`);

        const searchResults = await searchClient.search(searchQuery, {
            // TODO: re-enable after full re-index populates UserIds field
            filter: securityFilter,
            select: ["title", "content"],
            top: 3
        });

        const extractedTexts = [];
        let docCount = 0;

        for await (const result of searchResults.results) {
            docCount++;
            const doc = result.document;
            const title = doc.title || "Untitled";
            const content = doc.content || "";

            console.log(`[SEARCH] Hit ${docCount}: "${title}" (${content.length} chars)`);

            if (content.length > 0) {
                extractedTexts.push(`[Source: ${title}]\n${content.substring(0, 15000)}`);
            } else {
                extractedTexts.push(`[Source: ${title}]\n(Document found but content is empty)`);
            }
        }

        console.log(`[SEARCH] Total documents matched: ${docCount}`);

        if (extractedTexts.length === 0) {
            return { text: "No information found in your permitted policy documents." };
        }

        return { text: extractedTexts.join("\n\n---\n\n") };

    } catch (error) {
        console.error(`[SEARCH] Azure AI Search Error: ${error.message}`);
        console.error(`[SEARCH] Stack: ${error.stack}`);
        return { error: error.message };
    }
}

module.exports = { searchSharePoint };