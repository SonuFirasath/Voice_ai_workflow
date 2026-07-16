// lib/auth.js
// App-only Microsoft Graph token via client-credentials flow.

async function getAccessToken() {
  try {
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "client_credentials",
          client_id: process.env.CLIENT_ID,
          client_secret: process.env.CLIENT_SECRET,
          scope: "https://graph.microsoft.com/.default",
        }),
      },
    );

    const text = await tokenResponse.text();
    const data = text ? JSON.parse(text) : {};

    if (!data.access_token) {
      console.error("[AUTH] Failed to get access token:", data.error_description || text);
      return null;
    }

    return data.access_token;
  } catch (error) {
    console.error("[AUTH] Token request failed:", error.message);
    return null;
  }
}

module.exports = { getAccessToken };
