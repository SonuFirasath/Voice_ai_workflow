// src/functions/vapiWebhook.js

const { app } = require('@azure/functions');
const { getAccessToken }       = require('../lib/auth');
const { lookupCaller }         = require('../lib/identity');
const { buildAssistantConfig } = require('../lib/assistantConfig');
const { searchSharePoint }     = require('../lib/search');

app.http('vapiWebhook', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log('[WEBHOOK] Vapi Webhook triggered.');

        try {
            const body          = await request.json();
            const eventType     = body.message?.type;
            const callerNumber  = body.message?.call?.customer?.number;

            context.log(`[WEBHOOK] Event: ${eventType} | Caller: ${callerNumber || 'N/A'}`);

            // ── Authenticate with Microsoft Entra ID ──
            const accessToken = await getAccessToken();
            if (!accessToken) {
                context.log('[WEBHOOK] AUTH FAILED: no access token returned');
                return { status: 500, jsonBody: { error: "Authentication with Entra ID failed." } };
            }
            context.log('[WEBHOOK] Entra ID token acquired.');

            // ── Resolve caller identity ──
            let employeeName = "Employee";
            let callerId     = null;

            if (callerNumber) {
                const caller = await lookupCaller(accessToken, callerNumber);
                context.log(`[WEBHOOK] User lookup result: ${JSON.stringify(caller)}`);

                if (!caller) {
                    context.log(`[WEBHOOK] CALLER NOT FOUND in Entra ID: ${callerNumber}`);
                    // Don't block — allow the call to proceed but search will be limited
                    // For assistant-request, we still return config so the user hears the greeting
                    if (eventType === 'assistant-request') {
                        context.log('[WEBHOOK] Returning assistant config despite unidentified caller.');
                        const host = request.headers.get('x-forwarded-host') || request.headers.get('host');
                        const publicUrl = host ? `https://${host}/api/vapiWebhook` : undefined;
                        return { status: 200, jsonBody: buildAssistantConfig(employeeName, publicUrl) };
                    }
                    // For tool-calls from an unidentified caller, return a clear message
                    return {
                        status: 200,
                        jsonBody: {
                            results: [{
                                toolCallId: body.message?.toolCallList?.[0]?.id || body.message?.toolCalls?.[0]?.id || "unknown",
                                result: "I'm sorry, but I couldn't verify your identity in our directory. Please contact IT support to ensure your phone number is registered."
                            }]
                        }
                    };
                }

                employeeName = caller.displayName;
                callerId     = caller.id;
                context.log(`[WEBHOOK] Identified: ${employeeName} | ID: ${callerId}`);
            } else {
                context.log('[WEBHOOK] WARNING: No caller number in payload — skipping identity lookup.');
            }

            // ── EVENT A: assistant-request — return assistant configuration ──
            if (eventType === 'assistant-request') {
                context.log(`[WEBHOOK] Returning assistant config for: ${employeeName}`);

                const host = request.headers.get('x-forwarded-host') || request.headers.get('host');
                const publicUrl = host ? `https://${host}/api/vapiWebhook` : undefined;
                context.log(`[WEBHOOK] Server URL for tool callbacks: ${publicUrl || 'NOT RESOLVED'}`);

                return { status: 200, jsonBody: buildAssistantConfig(employeeName, publicUrl) };
            }

            // ── EVENT B: tool-calls — AI invoked a tool ──
            if (eventType === 'tool-calls') {
                const toolCall = body.message.toolCallList?.[0] || body.message.toolCalls?.[0];

                if (!toolCall) {
                    context.log(`[WEBHOOK] ERROR: No tool call found. Message keys: ${Object.keys(body.message)}`);
                    return {
                        status: 200,
                        jsonBody: {
                            results: [{
                                toolCallId: "unknown",
                                result: "I encountered an internal error. Please try asking your question again."
                            }]
                        }
                    };
                }

                const toolCallId = toolCall.id;
                context.log(`[WEBHOOK] Tool called: ${toolCall.function?.name} | toolCallId: ${toolCallId}`);

                if (toolCall.function?.name === 'search_sharepoint') {
                    let args;
                    try {
                        args = typeof toolCall.function.arguments === 'string'
                            ? JSON.parse(toolCall.function.arguments)
                            : toolCall.function.arguments;
                    } catch (parseErr) {
                        context.log(`[WEBHOOK] ERROR: Failed to parse tool arguments: ${parseErr.message}`);
                        return {
                            status: 200,
                            jsonBody: {
                                results: [{ toolCallId, result: "I had trouble understanding the request. Could you rephrase your question?" }]
                            }
                        };
                    }

                    context.log(`[WEBHOOK] Search query: "${args.query}" | Caller ID: "${callerId}"`);

                    if (!callerId) {
                        context.log('[WEBHOOK] CRITICAL: callerId is null at search time — identity lookup must have failed.');
                    }

                    const result = await searchSharePoint(args.query, callerId);

                    // result has either { text: "..." } or { error: "..." }
                    if (result.error) {
                        context.log(`[WEBHOOK] SEARCH ERROR: ${result.error}`);
                        return {
                            status: 200,
                            jsonBody: {
                                results: [{
                                    toolCallId,
                                    result: "I'm having trouble accessing the document system right now. Please try again in a moment."
                                }]
                            }
                        };
                    }

                    context.log(`[WEBHOOK] Search success: returning ${result.text.length} chars to AI.`);
                    return {
                        status: 200,
                        jsonBody: {
                            results: [{
                                toolCallId,
                                result: result.text
                            }]
                        }
                    };
                }

                // Unknown tool name
                context.log(`[WEBHOOK] Unknown tool: ${toolCall.function?.name}`);
                return {
                    status: 200,
                    jsonBody: {
                        results: [{
                            toolCallId,
                            result: "That function is not available."
                        }]
                    }
                };
            }

            // ── Unhandled event types ──
            context.log(`[WEBHOOK] Event type "${eventType}" not handled — ignoring.`);
            return { status: 200, jsonBody: { message: "Event ignored" } };

        } catch (error) {
            context.log(`[WEBHOOK] FATAL ERROR: ${error.message}`);
            context.log(`[WEBHOOK] Stack: ${error.stack}`);
            return { status: 500, jsonBody: { error: error.message } };
        }
    }
});