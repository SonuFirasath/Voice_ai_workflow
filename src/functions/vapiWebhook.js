// const { app } = require('@azure/functions');

// app.http('vapiWebhook', {
//     methods: ['POST'],
//     authLevel: 'anonymous', // Vapi needs unauthenticated access to trigger the webhook
//     handler: async (request, context) => {
//         context.log(`Webhook triggered by Vapi execution.`);

//         try {
//             // 1. Parse the incoming JSON body from Vapi
//             const body = await request.json();
            
//             // 2. Safely check if this is an 'assistant-request' event
//             const eventType = body.message?.type;
            
//             if (eventType === 'assistant-request') {
//                 // 3. Extract the caller's phone number
//                 const callerNumber = body.message.call?.customer?.number;
//                 context.log(`Incoming call detected from number: ${callerNumber}`);

//                 // 4. Temporary Step: Return a basic template assistant to Vapi 
//                 // This keeps the call alive while we test the webhook connection
//                 return {
//                     status: 200,
//                     headers: { 'Content-Type': 'application/json' },
//                     jsonBody: {
//                         assistant: {
//                             firstMessage: "Connection established with the Node.js backend. Ready for authentication step.",
//                             model: {
//                                 provider: "openai",
//                                 model: "gpt-4o"
//                             }
//                         }
//                     }
//                 };
//             }

//             // Fallback for other Vapi event types (like status updates or end-of-call logs)
//             return { status: 200, jsonBody: { message: "Event received successfully" } };

//         } catch (error) {
//             context.log(`Error processing webhook: ${error.message}`);
//             return {
//                 status: 500,
//                 jsonBody: { error: "Internal Server Error during payload processing" }
//             };
//         }
//     }
// });


const { app } = require('@azure/functions');

app.http('vapiWebhook', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log('Vapi Webhook triggered.');

        try {
            // --- STEP 1: Parse Vapi Payload ---
            const body = await request.json();
            const eventType = body.message?.type;

            if (eventType !== 'assistant-request') {
                return { status: 200, jsonBody: { message: "Event ignored" } };
            }

            const callerNumber = body.message.call?.customer?.number;
            context.log(`Incoming call from: ${callerNumber}`);

            if (!callerNumber) {
                return { status: 400, jsonBody: { error: "Missing caller number" } };
            }

            // --- STEP 2: Fetch Microsoft Entra ID Token ---
            // These environment variables will be stored in your local settings
            const tenantId = process.env.TENANT_ID;
            const clientId = process.env.CLIENT_ID;
            const clientSecret = process.env.CLIENT_SECRET;

            context.log('Requesting access token from Microsoft Entra ID...');
            
            const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
                body: new URLSearchParams({
                    grant_type: 'client_credentials',
                    client_id: clientId,
                    client_secret: clientSecret,
                    scope: 'https://graph.microsoft.com/.default'
                })
            });

            if (!tokenResponse.ok) {
                const errorData = await tokenResponse.text();
                context.log(`Token generation failed: ${errorData}`);
                return { status: 500, jsonBody: { error: "Authentication with Entra ID failed" } };
            }

            const tokenData = await tokenResponse.json();
            const accessToken = tokenData.access_token;
            context.log('Successfully obtained Entra ID access token.');

            // --- TEMPORARY RESPONSE FOR TESTING ---
            return {
                status: 200,
                headers: { 'Content-Type': 'application/json' },
                jsonBody: {
                    assistant: {
                        firstMessage: "Token generated successfully! Moving to validation step next.",
                        model: {
                            provider: "openai",
                            model: "gpt-4o"
                        }
                    }
                }
            };

        } catch (error) {
            context.log(`Error: ${error.message}`);
            return { status: 500, jsonBody: { error: error.message } };
        }
    }
});