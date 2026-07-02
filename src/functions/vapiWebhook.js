const { app } = require('@azure/functions');
const { getAccessToken }       = require('../lib/auth');
const { lookupCaller }         = require('../lib/identity');
const { buildAssistantConfig } = require('../lib/assistantConfig');
const { searchSharePoint }     = require('../lib/search');

app.http('vapiWebhook', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log('Vapi Webhook triggered.');

        try {
            const body          = await request.json();
            const eventType     = body.message?.type;
            const callerNumber  = body.message?.call?.customer?.number;

            context.log(`Event: ${eventType} | Caller: ${callerNumber || 'N/A'}`);

            // Authenticate with Microsoft Entra ID
            const accessToken = await getAccessToken();
            if (!accessToken) {
                context.log('AUTH FAILED: no access token returned');
                return { status: 500, jsonBody: { error: "Authentication with Entra ID failed." } };
            }

            // Resolve caller identity and role
            let employeeName = "Employee";
            let jobTitle     = "";

            if (callerNumber) {
                const caller = await lookupCaller(accessToken, callerNumber);
                context.log(`User lookup result: ${JSON.stringify(caller)}`);

                if (!caller) {
                    context.log(`CALLER NOT FOUND in Entra ID: ${callerNumber}`);
                    return { status: 200, jsonBody: { error: "Access denied. Number not found in Entra ID." } };
                }

                employeeName = caller.displayName;
                jobTitle     = caller.jobTitle;
                context.log(`Identified: ${employeeName} | Role: ${jobTitle}`);
            } else {
                context.log('WARNING: No caller number in payload — skipping identity lookup.');
            }

            // EVENT A: incoming call — return assistant configuration
            if (eventType === 'assistant-request') {
                context.log(`Returning assistant config for: ${employeeName}`);
                return { status: 200, jsonBody: buildAssistantConfig(employeeName) };
            }

            // EVENT B: AI invoked a tool
            if (eventType === 'tool-calls') {
                const toolCall = body.message.toolCallList?.[0] || body.message.toolCalls?.[0];

                if (!toolCall) {
                    context.log(`ERROR: No tool call found. Keys: ${Object.keys(body.message)}`);
                    return { status: 200, jsonBody: { message: "No tool call found" } };
                }

                context.log(`Tool called: ${toolCall.function.name}`);

                if (toolCall.function.name === 'search_sharepoint') {
                    const args = typeof toolCall.function.arguments === 'string'
                        ? JSON.parse(toolCall.function.arguments)
                        : toolCall.function.arguments;

                    context.log(`Search query: "${args.query}" | Role: "${jobTitle}"`);

                    const result = await searchSharePoint(accessToken, args.query, jobTitle);

                    if (result.error) {
                        context.log(`SEARCH API ERROR: ${JSON.stringify(result.error)}`);
                        return {
                            status: 200,
                            jsonBody: {
                                results: [{ toolCallId: toolCall.id, result: "The search system encountered an error. Please try again." }]
                            }
                        };
                    }

                    context.log(`Returning to AI: ${result.text.length} chars`);
                    return {
                        status: 200,
                        jsonBody: { results: [{ toolCallId: toolCall.id, result: result.text }] }
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
