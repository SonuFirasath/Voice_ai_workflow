// getFolderIds.js
const fs = require('fs');
const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
// (Add your tenantId, clientId, clientSecret logic here like the previous scripts)
// ... [Authentication logic same as previous scripts] ...

async function getIds() {
    // ... [Auth logic] ...
    // Fetch the site drive root
    const siteId = "chimeratechpvtltd.sharepoint.com,939496b9-b57f-440c-90c6-49115a93ed58,b535fe0e-a060-410d-8e0c-744aaa21e8be";
    const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children`, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    const data = await res.json();
    console.log(JSON.stringify(data.value, null, 2));
}