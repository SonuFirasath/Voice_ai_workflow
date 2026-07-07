// src/lib/identity.js

/**
 * Strips a phone number down to pure digits.
 * "+91 9876 543210"  →  "919876543210"
 * "(555) 123-4567"   →  "5551234567"
 */
function stripToDigits(phone) {
  return phone.replace(/\D/g, "");
}

/**
 * Builds an array of plausible phone-number formats that Entra ID
 * might store for a given caller number.
 *
 * VAPI always sends E.164 (e.g. "+919876543210" or "+15551234567").
 * Admins store numbers in many ways, so we try several variants.
 */
function buildPhoneVariants(rawNumber) {
  const variants = new Set();

  // 1. Exact as-received (e.g. "+919876543210")
  variants.add(rawNumber);

  const digits = stripToDigits(rawNumber); // e.g. "919876543210"

  // 2. All digits without "+" (e.g. "919876543210")
  variants.add(digits);

  // 3. With "+" prefix (e.g. "+919876543210") — may duplicate #1 but Set handles that
  variants.add(`+${digits}`);

  // Indian numbers: country code 91, 10-digit local number
  if (digits.startsWith("91") && digits.length === 12) {
    const local = digits.slice(2); // "9876543210"
    variants.add(local);
    variants.add(`+91${local}`);
    variants.add(`+91 ${local}`);
    variants.add(`091${local}`);
  }

  // US/CA numbers: country code 1, 10-digit local number
  if (digits.startsWith("1") && digits.length === 11) {
    const local = digits.slice(1); // "5551234567"
    variants.add(local);
    variants.add(`+1${local}`);
    variants.add(`+1 ${local}`);
    // Common US formatting: (555) 123-4567
    variants.add(
      `(${local.slice(0, 3)}) ${local.slice(3, 6)}-${local.slice(6)}`,
    );
    variants.add(`${local.slice(0, 3)}-${local.slice(3, 6)}-${local.slice(6)}`);
  }

  // If it's already 10 digits (no country code), also try with common country codes
  if (digits.length === 10) {
    variants.add(digits);
    variants.add(`+1${digits}`); // maybe US
    variants.add(`+91${digits}`); // maybe India
    // Common US formatting
    variants.add(
      `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`,
    );
    variants.add(
      `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6)}`,
    );
  }

  return [...variants];
}

/**
 * Build an OData filter expression that checks mobilePhone and businessPhones
 * against every plausible variant of the caller's number.
 */
function buildPhoneFilter(variants) {
  const conditions = variants.map((v) => {
    const escaped = v.replace(/'/g, "''"); // OData single-quote escaping
    return `(mobilePhone eq '${escaped}' or businessPhones/any(p:p eq '${escaped}'))`;
  });
  return conditions.join(" or ");
}

async function lookupCaller(accessToken, callerNumber) {
  console.log(`[IDENTITY] Looking up caller: ${callerNumber}`);

  const variants = buildPhoneVariants(callerNumber);
  console.log(
    `[IDENTITY] Trying ${variants.length} phone variants: ${JSON.stringify(variants)}`,
  );

  const filter = buildPhoneFilter(variants);
  const url = `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(filter)}&$select=id,displayName,jobTitle&$count=true`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ConsistencyLevel: "eventual",
    },
  });

  const data = await response.json();

  if (data.error) {
    console.error(`[IDENTITY] Graph API error: ${JSON.stringify(data.error)}`);
    return null;
  }

  if (!data.value || data.value.length === 0) {
    console.log(`[IDENTITY] No user found for any variant of: ${callerNumber}`);
    return null;
  }

  console.log(
    `[IDENTITY] Found user: ${data.value[0].displayName} (ID: ${data.value[0].id})`,
  );

  return {
    id: data.value[0].id,
    displayName: data.value[0].displayName || "Employee",
    jobTitle: (data.value[0].jobTitle || "").toLowerCase(),
  };
}

async function getUserGroupIds(accessToken, userId) {
  console.log(`[IDENTITY] Fetching group memberships for: ${userId}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${userId}/memberOf?$select=id&$top=100`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        ConsistencyLevel: "eventual",
      },
    }
  );
  const data = await res.json();
  if (data.error) {
    console.error(`[IDENTITY] Group lookup failed: ${JSON.stringify(data.error)}`);
    return [];
  }
  const ids = (data.value || []).map((g) => g.id);
  console.log(`[IDENTITY] User is member of ${ids.length} group(s)`);
  return ids;
}

module.exports = { lookupCaller, getUserGroupIds };
