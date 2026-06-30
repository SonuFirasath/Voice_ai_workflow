const { app } = require('@azure/functions');
const { PDFParse } = require('pdf-parse');

app.http('vapiWebhook', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log('Vapi Webhook triggered.');

        try {
            const body = await request.json();
            const eventType = body.message?.type;
            const callerNumber = body.message?.call?.customer?.number;
            
            context.log(`Event: ${eventType} | Caller: ${callerNumber || 'N/A'}`);

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

            if (!accessToken) {
                context.log(`AUTH FAILED: ${JSON.stringify(tokenData)}`);
                return { status: 500, jsonBody: { error: "Authentication with Entra ID failed." } };
            }

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
                context.log(`User lookup result: ${JSON.stringify(userData)}`);

                if (userData.value && userData.value.length > 0) {
                    employeeName = userData.value[0].displayName || "Employee";
                    jobTitle = (userData.value[0].jobTitle || "").toLowerCase();
                    context.log(`Identified: ${employeeName} | Role: ${jobTitle}`);
                } else {
                    context.log(`CALLER NOT FOUND in Entra ID: ${callerNumber}`);
                    return { status: 200, jsonBody: { error: "Access denied. Number not found in Entra ID." } };
                }
            } else {
                context.log('WARNING: No caller number in payload — skipping identity lookup.');
            }

            // --- EVENT A: INCOMING CALL (SETUP ASSISTANT) ---
            if (eventType === 'assistant-request') {
                context.log(`Returning assistant config for: ${employeeName}`);
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
                const toolCall = body.message.toolCallList?.[0] || body.message.toolCalls?.[0];
                
                if (!toolCall) {
                    context.log(`ERROR: No tool call found in payload. Keys: ${Object.keys(body.message)}`);
                    return { status: 200, jsonBody: { message: "No tool call found" } };
                }

                context.log(`Tool called: ${toolCall.function.name}`);

                if (toolCall.function.name === 'search_sharepoint') {
                    const args = typeof toolCall.function.arguments === 'string'
                        ? JSON.parse(toolCall.function.arguments)
                        : toolCall.function.arguments;
                    const searchQuery = args.query;
                    
                    context.log(`Search query from AI: "${searchQuery}"`);
                    context.log(`Caller jobTitle: "${jobTitle}"`);
                    
                    // RBAC: Build the allowed folders list based on the caller's job title
                    let allowedPaths = `path:"${process.env.SP_GENERAL_FOLDER}/*"`;
                    
                    if (jobTitle.includes("manager")) {
                        allowedPaths += ` OR path:"${process.env.SP_MANAGER_FOLDER}/*"`;
                    }
                    if (jobTitle.includes("sales")) {
                        allowedPaths += ` OR path:"${process.env.SP_SALES_FOLDER}/*"`;
                    }

                    const fullQuery = `${searchQuery} AND (${allowedPaths})`;
                    context.log(`Full KQL query: ${fullQuery}`);

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
                                query: { queryString: fullQuery },
                                region: "IND"
                            }]
                        })
                    });

                    const searchData = await searchResponse.json();

                    // Check for API-level errors (e.g., missing region, permission issues)
                    if (searchData.error) {
                        context.log(`SEARCH API ERROR: ${JSON.stringify(searchData.error)}`);
                        return {
                            status: 200,
                            jsonBody: {
                                results: [{
                                    toolCallId: toolCall.id,
                                    result: "The search system encountered an error. Please try again."
                                }]
                            }
                        };
                    }

                    let resultText = "No information found in the allowed policy documents.";

                    const hits = searchData.value?.[0]?.hitsContainers?.[0]?.hits || [];
                    const totalResults = searchData.value?.[0]?.hitsContainers?.[0]?.total || 0;
                    
                    context.log(`Search results: ${hits.length} hits (total: ${totalResults})`);

                    if (hits.length > 0) {
                        // Download and extract full text from matched PDFs (limit to top 3)
                        const extractedTexts = [];
                        const maxFiles = Math.min(hits.length, 3);

                        for (let i = 0; i < maxFiles; i++) {
                            const hit = hits[i];
                            const fileName = hit.resource?.name || "Unknown";
                            const resourceId = hit.resource?.id;
                            const driveId = hit.resource?.parentReference?.driveId;

                            context.log(`  Downloading [${i + 1}]: ${fileName}`);

                            if (!driveId || !resourceId) {
                                context.log(`  Skipping — missing driveId or resourceId`);
                                // Fall back to search snippet
                                if (hit.summary) extractedTexts.push(`[Source: ${fileName}]\n${hit.summary}`);
                                continue;
                            }

                            try {
                                const downloadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${resourceId}/content`;
                                const downloadRes = await fetch(downloadUrl, {
                                    headers: { 'Authorization': `Bearer ${accessToken}` }
                                });

                                if (!downloadRes.ok) {
                                    context.log(`  Download failed: ${downloadRes.status}`);
                                    if (hit.summary) extractedTexts.push(`[Source: ${fileName}]\n${hit.summary}`);
                                    continue;
                                }

                                const fileBuffer = Buffer.from(await downloadRes.arrayBuffer());
                                context.log(`  Downloaded: ${fileBuffer.length} bytes`);

                                if (fileName.toLowerCase().endsWith('.pdf')) {
                                    const parser = new PDFParse({ data: fileBuffer });
                                    const pdfData = await parser.getText();
                                    await parser.destroy();
                                    
                                    const fullText = pdfData.text;
                                    const snippetText = fullText.substring(0, 40000);
                                    
                                    // Clean up the text for voice reading
                                    const cleanText = snippetText
                                        .replace(/--\s*\d+\s*of\s*\d+\s*--/g, '') // Remove page numbers like "-- 5 of 44 --"
                                        .replace(/CT\/EH\/\d+\.\d+/g, '')         // Remove document codes like "CT/EH/3.1"
                                        .replace(/--/g, '')                       // Remove stray double dashes
                                        .replace(/\.{3,}/g, '.')                  // Condense TOC dots "......" into a single "."
                                        .replace(/\s+/g, ' ')                     // Compress newlines and multiple spaces
                                        .trim();
                                        
                                    context.log(`  Extracted first ${snippetText.length} chars (to bypass TOC), cleaned to ${cleanText.length} chars.`);
                                    
                                    extractedTexts.push(`[Source: ${fileName}]\n${cleanText}`);
                                } else {
                                    const rawText = fileBuffer.toString('utf8').substring(0, 15000);
                                    extractedTexts.push(`[Source: ${fileName}]\n${rawText}`);
                                }
                            } catch (dlErr) {
                                context.log(`  Error processing ${fileName}: ${dlErr.message}`);
                                if (hit.summary) extractedTexts.push(`[Source: ${fileName}]\n${hit.summary}`);
                            }
                        }

                        if (extractedTexts.length > 0) {
                            resultText = extractedTexts.join("\n\n---\n\n");
                        }
                    }

                    context.log(`Returning to AI: ${resultText.length} chars`);

                    return {
                        status: 200,
                        jsonBody: {
                            results: [{
                                toolCallId: toolCall.id,
                                result: resultText
                            }]
                        }
                    };
                }
            }

            context.log(`Event type "${eventType}" not handled — ignoring.`);
            return { status: 200, jsonBody: { message: "Event ignored" } };

        } catch (error) {
            context.log(`FATAL ERROR: ${error.message}`);
            context.log(`Stack: ${error.stack}`);
            return { status: 500, jsonBody: { error: error.message } };
        }
    }
});