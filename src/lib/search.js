const { downloadAndExtract } = require('./fileExtractor');

function buildAllowedPaths(jobTitle) {
    let paths = `path:"${process.env.SP_GENERAL_FOLDER}/*"`;
    if (jobTitle.includes("manager")) paths += ` OR path:"${process.env.SP_MANAGER_FOLDER}/*"`;
    if (jobTitle.includes("sales"))   paths += ` OR path:"${process.env.SP_SALES_FOLDER}/*"`;
    return paths;
}

async function searchSharePoint(accessToken, searchQuery, jobTitle) {
    const allowedPaths = buildAllowedPaths(jobTitle);
    const fullQuery = `${searchQuery} AND (${allowedPaths})`;

    const searchResponse = await fetch('https://graph.microsoft.com/v1.0/search/query', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            requests: [{
                entityTypes: ["driveItem"],
                query: { queryString: fullQuery },
                region: "IND"
            }]
        })
    });

    const searchData = await searchResponse.json();

    if (searchData.error) {
        return { error: searchData.error };
    }

    const hits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];

    if (hits.length === 0) {
        return { text: "No information found in the allowed policy documents." };
    }

    const extractedTexts = [];
    const maxFiles = Math.min(hits.length, 3);

    for (let i = 0; i < maxFiles; i++) {
        try {
            const text = await downloadAndExtract(accessToken, hits[i]);
            if (text) extractedTexts.push(text);
        } catch (err) {
            const fallback = hits[i].summary;
            if (fallback) extractedTexts.push(`[Source: ${hits[i].resource?.name || "Unknown"}]\n${fallback}`);
        }
    }

    const text = extractedTexts.length > 0
        ? extractedTexts.join("\n\n---\n\n")
        : "No information found in the allowed policy documents.";

    return { text };
}

module.exports = { searchSharePoint };
