// src/lib/assistantConfig.js

function buildAssistantConfig(employeeName, serverUrl) {
  // Guard: if serverUrl wasn't resolved, log it clearly
  if (!serverUrl) {
    console.error(
      "[CONFIG] WARNING: serverUrl is undefined — VAPI tool calls will fail to route back.",
    );
  }

  const toolConfig = {
    type: "function",
    function: {
      name: "search_sharepoint",
      description:
        "Searches the permitted SharePoint policy documents for company policies and information. Use this tool whenever the user asks a question about company policies, procedures, benefits, or any work-related information.",
      parameters: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description:
              "The search keywords extracted from the user's question.",
          },
        },
        required: ["query"],
      },
    },
    async: false,
  };

  if (serverUrl) {
    toolConfig.server = {
      url: serverUrl,
    };
  }

  return {
    assistant: {
      name: "Enterprise Agent",
      firstMessage: `Hello, ${employeeName}. How can I help you today?`,
      model: {
        provider: "openai",
        model: "gpt-4o",
        messages: [
          {
            role: "system",
            content: `You are a corporate policy assistant for ${employeeName}. Your ONLY job is to answer questions using the search_sharepoint tool.

RULES:
1. For EVERY user question, call the search_sharepoint tool first.
2. Your answer must be built exclusively from the exact text returned by the tool. Never use your training knowledge to add context, explain terms, or fill in gaps — even if you recognise a tool or system mentioned in the documents.
3. If the tool returns no information, say: "Sorry, I couldn't find an answer to that in your permitted documents."
4. If a tool, system, or application is named in the documents but the documents do not provide step-by-step instructions for using it, say: "Sorry, I don't have detailed instructions for that in your permitted documents."
5. Keep answers conversational and brief — this is a phone call, not an essay.`,
          },
        ],
        tools: [toolConfig],
      },
    },
  };
}

module.exports = { buildAssistantConfig };
