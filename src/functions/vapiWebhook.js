// src/functions/vapiWebhook.js

const { app } = require("@azure/functions");
const { getAccessToken } = require("../lib/auth");
const { lookupCaller } = require("../lib/identity");
const { searchSharePoint } = require("../lib/search");
const { checkAvailability, bookAppointment } = require("../lib/calendar");

app.http("vapiWebhook", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const eventType = body.message?.type;
      const callerNumber = body.message?.call?.customer?.number;

      // ── 1. DYNAMIC GREETING (Inbound Call Starts) ──
      if (eventType === "assistant-request") {
        let employeeName = "there";

        const accessToken = await getAccessToken();
        if (accessToken && callerNumber) {
          const caller = await lookupCaller(accessToken, callerNumber);
          if (caller) {
            employeeName = caller.displayName;
          }
        }

        context.log(
          `[LOW-CODE GATEWAY] Generating dynamic greeting for: ${employeeName}`,
        );

        // The server returns ONLY the dynamic first message. Vapi merges this with the Dashboard setup!
        return {
          status: 200,
          jsonBody: {
            assistant: {
              firstMessage: `Authentication successful. Welcome, ${employeeName}. How can assistance be provided today?`,
            },
          },
        };
      }

      // ── 2. TOOL EXECUTION (AI searches Microsoft Graph) ──
      if (eventType === "tool-calls") {
        const toolName = toolCall?.function?.name;
        let searchResult;

        if (toolName === "search_sharepoint") {
          searchResult = await searchSharePoint(args.query, null);
        } else if (toolName === "check_availability") {
          searchResult = await checkAvailability(args.date);
        } else if (toolName === "book_appointment") {
          searchResult = await bookAppointment(args.time, args.topic);
        } else {
          return {
            status: 200,
            jsonBody: {
              results: [
                { toolCallId: toolCall?.id, result: "Tool not available." },
              ],
            },
          };
        }

        if (searchResult.error) {
          return {
            status: 200,
            jsonBody: {
              results: [
                {
                  toolCallId: toolCall.id,
                  result: "Calendar access error. Please try again.",
                },
              ],
            },
          };
        }

        return {
          status: 200,
          jsonBody: {
            results: [{ toolCallId: toolCall.id, result: searchResult.text }],
          },
        };
      }

      return { status: 200, jsonBody: { message: "Event ignored" } };
    } catch (error) {
      context.log(`[LOW-CODE GATEWAY ERROR] ${error.message}`);
      return { status: 500, jsonBody: { error: error.message } };
    }
  },
});
