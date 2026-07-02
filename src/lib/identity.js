async function lookupCaller(accessToken, callerNumber) {
    const encodedNumber = encodeURIComponent(callerNumber);
    const url = `https://graph.microsoft.com/v1.0/users?$filter=mobilePhone eq '${encodedNumber}' or businessPhones/any(p:p eq '${encodedNumber}')&$select=displayName,jobTitle&$count=true`;

    const response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'ConsistencyLevel': 'eventual'
        }
    });

    const data = await response.json();

    if (!data.value || data.value.length === 0) return null;

    return {
        displayName: data.value[0].displayName || "Employee",
        jobTitle: (data.value[0].jobTitle || "").toLowerCase()
    };
}

module.exports = { lookupCaller };
