const fs = require('fs');
const { PDFParse } = require('pdf-parse');
const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;
(async () => {
    const tokenRes = await fetch('https://login.microsoftonline.com/'+tenantId+'/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({ grant_type: 'client_credentials', client_id: clientId, client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default' })
    });
    const token = (await tokenRes.json()).access_token;
    
    const allowedPaths = `path:"${localSettings.Values.SP_GENERAL_FOLDER}/*" OR path:"${localSettings.Values.SP_MANAGER_FOLDER}/*"`;
    const fullQuery = `What is our brand promise AND (${allowedPaths})`;
    const searchRes = await fetch('https://graph.microsoft.com/v1.0/search/query', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ requests: [{ entityTypes: ['driveItem'], query: { queryString: fullQuery }, region: 'IND' }] })
    });
    const searchData = await searchRes.json();
    const hit = searchData.value[0].hitsContainers[0].hits[0];
    
    const downloadUrl = `https://graph.microsoft.com/v1.0/drives/${hit.resource.parentReference.driveId}/items/${hit.resource.id}/content`;
    const downloadRes = await fetch(downloadUrl, { headers: { 'Authorization': 'Bearer ' + token } });
    const fileBuffer = Buffer.from(await downloadRes.arrayBuffer());
    
    const parser = new PDFParse({ data: fileBuffer });
    const pdfData = await parser.getText();
    await parser.destroy();
    
    console.log('Total text length:', pdfData.text.length);
    const text = pdfData.text.toLowerCase();
    
    // Find all occurrences
    function getAllIndices(str, val) {
        let indices = [];
        let i = -1;
        while ((i = str.indexOf(val, i + 1)) != -1) {
            indices.push(i);
        }
        return indices;
    }

    const brandIndices = getAllIndices(text, 'brand promise');
    console.log('Indices of "brand promise":', brandIndices);
    
    const ceoIndices = getAllIndices(text, 'message from the ceo');
    console.log('Indices of "message from the ceo":', ceoIndices);
    
    if (brandIndices.length > 1) {
        let idx = brandIndices[1]; // Usually the first one is the Table of Contents
        console.log('\n--- Text snippet near SECOND "brand promise" ---');
        console.log(pdfData.text.substring(Math.max(0, idx - 100), idx + 500));
    }
    
    if (ceoIndices.length > 1) {
        let idx = ceoIndices[1]; // Usually the first one is the Table of Contents
        console.log('\n--- Text snippet near SECOND "message from the ceo" ---');
        console.log(pdfData.text.substring(Math.max(0, idx - 100), idx + 500));
    }
})();
