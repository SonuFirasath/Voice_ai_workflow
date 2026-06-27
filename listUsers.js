// listUsers.js
const fs = require('fs');

// Read the credentials directly from your local.settings.json file
const localSettings = JSON.parse(fs.readFileSync('./local.settings.json', 'utf8'));
const tenantId = localSettings.Values.TENANT_ID;
const clientId = localSettings.Values.CLIENT_ID;
const clientSecret = localSettings.Values.CLIENT_SECRET;

async function listAllEmployees() {
    console.log("Fetching Entra ID Access Token...");

    // 1. Get the Token
    const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            grant_type: 'client_credentials',
            client_id: clientId,
            client_secret: clientSecret,
            scope: 'https://graph.microsoft.com/.default'
        })
    });

    if (!tokenResponse.ok) {
        console.error("Auth Failed:", await tokenResponse.text());
        return;
    }
    const tokenData = await tokenResponse.json();
    const accessToken = tokenData.access_token;
    console.log("Token secured. Fetching users...\n");

    // 2. Fetch the Users
    // Using $select to pull specific useful fields. 
    const graphUrl = `https://graph.microsoft.com/v1.0/users?$select=displayName,userPrincipalName,mobilePhone,businessPhones`;

    const graphResponse = await fetch(graphUrl, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        }
    });

    if (!graphResponse.ok) {
        console.error("Graph API Error:", await graphResponse.text());
        return;
    }

    const graphData = await graphResponse.json();
    const users = graphData.value;

    // 3. Log them beautifully to the console
    console.log("================ ACTIVE ENTRA ID EMPLOYEES ================");
    users.forEach((user, index) => {
        const mobile = user.mobilePhone || "N/A";
        const business = (user.businessPhones && user.businessPhones.length > 0) ? user.businessPhones[0] : "N/A";
        
        console.log(`[${index + 1}] Name: ${user.displayName}`);
        console.log(`    Email:  ${user.userPrincipalName}`);
        console.log(`    Mobile: ${mobile} | Business: ${business}`);
        console.log("-----------------------------------------------------------");
    });
    console.log(`Total Employees Found: ${users.length}`);
}

listAllEmployees();