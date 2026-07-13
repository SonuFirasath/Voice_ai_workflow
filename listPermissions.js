// listPermissions.js — shows who has access to each SharePoint file
// Run with: node listPermissions.js

const fs = require("fs");
const localSettings = JSON.parse(
  fs.readFileSync("./local.settings.json", "utf8"),
);
Object.assign(process.env, localSettings.Values);

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SP_SITE_URL = process.env.SP_SITE_URL;

async function getAccessToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default",
      }),
    },
  );
  const data = await res.json();
  if (!data.access_token)
    throw new Error(
      `Auth failed: ${data.error_description || JSON.stringify(data)}`,
    );
  return data.access_token;
}

async function graphGet(token, url) {
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph ${res.status} @ ${url}\n${text}`);
  }
  return res.json();
}

async function listFilesRecursive(token, driveId, itemId, folderPath, results) {
  const url =
    itemId === "root"
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children`;

  let nextUrl = url;
  while (nextUrl) {
    const data = await graphGet(token, nextUrl);
    for (const item of data.value || []) {
      if (item.folder) {
        await listFilesRecursive(
          token,
          driveId,
          item.id,
          `${folderPath}/${item.name}`,
          results,
        );
      } else if (item.file) {
        results.push({ id: item.id, name: item.name, path: folderPath });
      }
    }
    nextUrl = data["@odata.nextLink"] || null;
  }
}

function formatPerm(perm) {
  const roles = (perm.roles || []).join(", ") || "read";

  // Direct user grant
  if (perm.grantedToV2?.user) {
    const u = perm.grantedToV2.user;
    return `   👤 ${u.displayName || u.email || u.id} — ${roles}`;
  }
  // Group grant
  if (perm.grantedToV2?.group) {
    const g = perm.grantedToV2.group;
    return `   👥 GROUP: ${g.displayName || g.id} — ${roles}`;
  }
  // Multiple identities (e.g. sharing links with specific people)
  if (perm.grantedToIdentitiesV2?.length) {
    return perm.grantedToIdentitiesV2
      .map((identity) => {
        if (identity.user)
          return `   👤 ${identity.user.displayName || identity.user.email} — ${roles}`;
        if (identity.group)
          return `   👥 GROUP: ${identity.group.displayName} — ${roles}`;
        return `   ? unknown identity — ${roles}`;
      })
      .join("\n");
  }
  // Sharing link
  if (perm.link) {
    const scope = perm.link.scope || "unknown scope";
    const type = perm.link.type || "unknown type";
    return `   🔗 Sharing link (${type}, ${scope}) — ${roles}`;
  }
  // Inherited with no grantee info
  if (perm.inheritedFrom) {
    return `   ↑  Inherited from: ${perm.inheritedFrom.path} — ${roles}`;
  }
  return `   ? Unresolved permission entry — ${roles}`;
}

async function run() {
  console.log("Authenticating with Entra ID...");
  const token = await getAccessToken();
  console.log("✅ Token acquired\n");

  // Resolve the SharePoint site
  const siteUrl = new URL(SP_SITE_URL);
  const hostname = siteUrl.hostname; // e.g. chimera.sharepoint.com
  const sitePath = siteUrl.pathname; // e.g. /sites/IT

  console.log(`Looking up site: ${SP_SITE_URL}`);
  const site = await graphGet(
    token,
    `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`,
  );
  console.log(`✅ Site: "${site.displayName}" (ID: ${site.id})\n`);

  // List drives (document libraries)
  const drivesData = await graphGet(
    token,
    `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`,
  );
  const drives = drivesData.value || [];
  console.log(`Document libraries found (${drives.length}):`);
  drives.forEach((d) => console.log(`  • ${d.name}  (ID: ${d.id})`));

  const defaultDrive = drives.find((d) => d.name === "Documents") || drives[0];
  if (!defaultDrive) throw new Error("No document library found in this site.");
  console.log(`\nUsing library: "${defaultDrive.name}"\n`);

  // Enumerate all files
  console.log("Scanning files...");
  const allFiles = [];
  await listFilesRecursive(token, defaultDrive.id, "root", "", allFiles);
  console.log(`Found ${allFiles.length} file(s).\n`);

  if (allFiles.length === 0) {
    console.log(
      "No files found. Check that SP_SITE_URL points to the correct site and the app has Sites.Read.All permission.",
    );
    return;
  }

  console.log("=".repeat(70));
  console.log("  SHAREPOINT PERMISSION REPORT");
  console.log("=".repeat(70));

  for (const file of allFiles) {
    const displayPath = `${file.path || "/"}/${file.name}`;
    console.log(`\n📄 ${displayPath}`);

    try {
      const permsData = await graphGet(
        token,
        `https://graph.microsoft.com/v1.0/drives/${defaultDrive.id}/items/${file.id}/permissions`,
      );

      const perms = permsData.value || [];
      if (perms.length === 0) {
        console.log(
          "   (no explicit permissions — fully inherits from parent)",
        );
        continue;
      }

      for (const perm of perms) {
        const inherited = perm.inheritedFrom
          ? ` [inherited from ${perm.inheritedFrom.path}]`
          : "";
        const line = formatPerm(perm);
        console.log(line + inherited);
      }
    } catch (e) {
      console.log(
        `   ⚠️  Could not fetch permissions: ${e.message.split("\n")[0]}`,
      );
    }
  }

  console.log(`\n${"=".repeat(70)}`);
  console.log(`Report complete — ${allFiles.length} files checked.`);
  console.log("=".repeat(70));
}

run().catch((err) => {
  console.error("\nFATAL:", err.message);
  if (err.message.includes("403") || err.message.includes("Forbidden")) {
    console.error(
      '\n→ The app registration is missing "Sites.Read.All" (or Sites.FullControl.All) in Microsoft Graph API permissions.',
    );
    console.error(
      "  Go to: Azure Portal → App Registrations → your app → API Permissions → Add permission → Microsoft Graph → Application → Sites.Read.All → Grant admin consent",
    );
  }
});
