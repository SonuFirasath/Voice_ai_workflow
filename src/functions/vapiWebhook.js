const { app } = require("@azure/functions");

app.http("vapiWebhook", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    context.log("Vapi Webhook triggered.");

    try {
      // STEP 1: Parse Payload
      const body = await request.json();
      const eventType = body.message?.type;

      if (eventType !== "assistant-request") {
        return { status: 200, jsonBody: { message: "Event ignored" } };
      }

      const callerNumber = body.message.call?.customer?.number;
      context.log(`Incoming call from: ${callerNumber}`);

      if (!callerNumber) {
        return { status: 400, jsonBody: { error: "Missing caller number" } };
      }

      // STEP 2: Fetch Entra ID Token
      const tenantId = process.env.TENANT_ID;
      const clientId = process.env.CLIENT_ID;
      const clientSecret = process.env.CLIENT_SECRET;

      const tokenResponse = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            grant_type: "client_credentials",
            client_id: clientId,
            client_secret: clientSecret,
            scope: "https://graph.microsoft.com/.default",
          }),
        },
      );

      if (!tokenResponse.ok) {
        return { status: 500, jsonBody: { error: "Entra ID auth failed" } };
      }

      const tokenData = await tokenResponse.json();
      const accessToken = tokenData.access_token;

      // STEP 3: Query Microsoft Graph API
      const encodedNumber = encodeURIComponent(callerNumber);

      // 1. Add &$count=true to the end of the URL
      const graphQueryUrl = `https://graph.microsoft.com/v1.0/users?$filter=mobilePhone eq '${encodedNumber}' or businessPhones/any(p:p eq '${encodedNumber}')&$select=displayName,jobTitle&$count=true`;

      context.log(`Querying Graph API for: ${callerNumber}`);

      const graphResponse = await fetch(graphQueryUrl, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
          ConsistencyLevel: "eventual", // 2. Required header for advanced filtering
        },
      });

      if (!graphResponse.ok) {
        const graphError = await graphResponse.text();
        context.log(`Graph API error: ${graphError}`);
        return { status: 500, jsonBody: { error: "Graph API query failed" } };
      }

      const graphData = await graphResponse.json();
      const users = graphData.value;

      // STEP 4: Return Dynamic Vapi Response
      if (users && users.length > 0) {
        const employeeName = users[0].displayName || "Employee";
        context.log(`Authorized caller identified: ${employeeName}`);

        // Proceed with a transient assistant configuration
        return {
          status: 200,
          headers: { "Content-Type": "application/json" },
          jsonBody: {
            assistant: {
              name: "Employee IT Portal",
              firstMessage: `Authentication successful. Welcome to the corporate network, ${employeeName}. How can assistance be provided today?`,
              model: {
                provider: "openai",
                model: "gpt-4o",
                messages: [
                  {
                    role: "system",
                    content: `The assistant is an internal IT support agent for ${employeeName}.`,
                  },
                ],
              },
            },
          },
        };
      } else {
        context.log(`Unauthorized caller: ${callerNumber}. Rejecting call.`);

        // Reject spam callers with a spoken error message
        return {
          status: 200,
          headers: { "Content-Type": "application/json" },
          jsonBody: {
            error:
              "Access denied. The phone number is not recognized in the directory. Goodbye.",
          },
        };
      }
    } catch (error) {
      context.log(`Error: ${error.message}`);
      return { status: 500, jsonBody: { error: error.message } };
    }
  },
});
