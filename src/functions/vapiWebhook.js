const { app } = require('@azure/functions');

app.http('vapiWebhook', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log('Vapi Webhook triggered.');

        try {
            const body = await request.json();
            const eventType = body.message?.type;
            const callerNumber = body.message.call?.customer?.number;
            
            // --- 1. AUTHENTICATE WITH MICROSOFT ENTRA ID ---
            const tokenResponse = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    grant_type: 'client_credentials',
                    client_id: process.env.CLIENT_ID,
                    client_secret: process.env.CLIENT_SECRET,
                    scope: 'https://graph.microsoft.com/.default'
                })
            });

            const tokenData = await tokenResponse.json();
            const accessToken = tokenData.access_token;

            // --- 2. FETCH CALLER IDENTITY & ROLE ---
            let employeeName = "Employee";
            let jobTitle = "";

            if (callerNumber) {
                const encodedNumber = encodeURIComponent(callerNumber);
                const userQueryUrl = `https://graph.microsoft.com/v1.0/users?$filter=mobilePhone eq '${encodedNumber}' or businessPhones/any(p:p eq '${encodedNumber}')&$select=displayName,jobTitle&$count=true`;

                const userResponse = await fetch(userQueryUrl, {
                    method: 'GET',
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'ConsistencyLevel': 'eventual'
                    }
                });

                const userData = await userResponse.json();
                if (userData.value && userData.value.length > 0) {
                    employeeName = userData.value[0].displayName || "Employee";
                    jobTitle = (userData.value[0].jobTitle || "").toLowerCase();
                } else {
                    return { status: 200, jsonBody: { error: "Access denied. Number not found in Entra ID." } };
                }
            }

            // --- EVENT A: INCOMING CALL (SETUP ASSISTANT) ---
            if (eventType === 'assistant-request') {
                return {
                    status: 200,
                    jsonBody: {
                        assistant: {
                            name: "Enterprise Agent",
                            firstMessage: `Authentication successful. Welcome, ${employeeName}. I can answer questions about your company policies. How can I help you today?`,
                            model: {
                                provider: "openai",
                                model: "gpt-4o",
                                messages: [
                                    {
                                        role: "system",
                                        content: "You are a corporate assistant. Use the search_sharepoint tool to answer questions. If the tool returns no information, state: 'Sorry, I am unable to find an answer to your question in your permitted files.' Never invent answers."
                                    }
                                ],
                                tools: [{
                                    type: "function",
                                    function: {
                                        name: "search_sharepoint",
                                        description: "Searches the permitted SharePoint policy documents.",
                                        parameters: {
                                            type: "object",
                                            properties: { query: { type: "string", description: "The search keywords." } },
                                            required: ["query"]
                                        }
                                    }
                                }]
                            }
                        }
                    }
                };
            }

            // --- EVENT B: AI USES THE SEARCH TOOL ---
            if (eventType === 'tool-calls') {
                const toolCall = body.message.toolCalls[0];
                
                if (toolCall.function.name === 'search_sharepoint') {
                    const searchQuery = JSON.parse(toolCall.function.arguments).query;
                    
                    // RBAC: Build the allowed folders list based on the caller's job title
                    let allowedPaths = `path:"${process.env.SP_GENERAL_FOLDER}"`;
                    
                    if (jobTitle.includes("manager")) {
                        allowedPaths += ` OR path:"${process.env.SP_MANAGER_FOLDER}"`;
                    }
                    if (jobTitle.includes("sales")) {
                        allowedPaths += ` OR path:"${process.env.SP_SALES_FOLDER}"`;
                    }

                    // Execute Microsoft Graph Search
                    const searchApiUrl = 'https://graph.microsoft.com/v1.0/search/query';
                    const searchResponse = await fetch(searchApiUrl, {
                        method: 'POST',
                        headers: {
                            'Authorization': `Bearer ${accessToken}`,
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({
                            requests: [{
                                entityTypes: ["driveItem"],
                                query: { queryString: `${searchQuery} AND (${allowedPaths})` }
                            }]
                        })
                    });

                    const searchData = await searchResponse.json();
                    let snippets = "No information found in the allowed policy documents.";

                    // Extract the text summaries if Microsoft found a match
                    if (searchData.value && searchData.value[0].hitsContainers[0].hits) {
                        const hits = searchData.value[0].hitsContainers[0].hits;
                        snippets = hits.map(hit => hit.summary).join("\n\n");
                    }

                    return {
                        status: 200,
                        jsonBody: {
                            results: [{
                                toolCallId: toolCall.id,
                                result: snippets
                            }]
                        }
                    };
                }
            }

            return { status: 200, jsonBody: { message: "Event ignored" } };

        } catch (error) {
            context.log(`Error: ${error.message}`);
            return { status: 500, jsonBody: { error: error.message } };
        }
    }
});