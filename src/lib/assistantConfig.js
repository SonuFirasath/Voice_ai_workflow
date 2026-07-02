function buildAssistantConfig(employeeName) {
    return {
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
                tools: [
                    {
                        type: "function",
                        function: {
                            name: "search_sharepoint",
                            description: "Searches the permitted SharePoint policy documents.",
                            parameters: {
                                type: "object",
                                properties: {
                                    query: { type: "string", description: "The search keywords." }
                                },
                                required: ["query"]
                            }
                        }
                    }
                ]
            }
        }
    };
}

module.exports = { buildAssistantConfig };
