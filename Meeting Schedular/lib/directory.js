// lib/directory.js
// Resolves attendee names spoken by the caller into Entra ID email addresses,
// the same directory-lookup pattern the main app uses for caller identity.

async function findUsersByName(accessToken, name) {
  const escaped = name.replace(/'/g, "''");
  const filter = `startswith(displayName,'${escaped}')`;
  const url = `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(filter)}&$select=id,displayName,mail&$count=true`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ConsistencyLevel: "eventual",
    },
  });

  const data = await response.json();

  if (data.error) {
    console.error(`[DIRECTORY] Graph API error: ${JSON.stringify(data.error)}`);
    return [];
  }

  return data.value || [];
}

async function resolveAttendee(accessToken, name) {
  const matches = (await findUsersByName(accessToken, name)).filter((u) => u.mail);

  if (matches.length === 1) {
    return { name, email: matches[0].mail, displayName: matches[0].displayName };
  }

  if (matches.length === 0) {
    return { name, email: null, error: `No directory match found for "${name}".` };
  }

  return {
    name,
    email: null,
    error: `Multiple people named "${name}" found; need a more specific name or an email address.`,
    candidates: matches.map((u) => u.displayName),
  };
}

async function resolveAttendees(accessToken, names) {
  const results = await Promise.all(names.map((name) => resolveAttendee(accessToken, name)));
  return {
    resolved: results.filter((r) => r.email),
    unresolved: results.filter((r) => !r.email),
  };
}

module.exports = { resolveAttendee, resolveAttendees };
