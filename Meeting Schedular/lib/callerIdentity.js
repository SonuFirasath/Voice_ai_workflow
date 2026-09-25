// lib/callerIdentity.js
// Resolves the caller's phone number to an Entra ID user (same phone-variant
// matching technique as the main app's identity.js), then reads their MPIN
// from the "Mpin" custom security attribute for PIN verification.

function stripToDigits(phone) {
  return phone.replace(/\D/g, "");
}

function buildPhoneVariants(rawNumber) {
  const variants = new Set();
  variants.add(rawNumber);

  const digits = stripToDigits(rawNumber);
  variants.add(digits);
  variants.add(`+${digits}`);

  if (digits.startsWith("91") && digits.length === 12) {
    const local = digits.slice(2);
    variants.add(local);
    variants.add(`+91${local}`);
    variants.add(`+91 ${local}`);
    variants.add(`091${local}`);
  }

  if (digits.startsWith("1") && digits.length === 11) {
    const local = digits.slice(1);
    variants.add(local);
    variants.add(`+1${local}`);
    variants.add(`+1 ${local}`);
    variants.add(`(${local.slice(0, 3)}) ${local.slice(3, 6)}-${local.slice(6)}`);
    variants.add(`${local.slice(0, 3)}-${local.slice(3, 6)}-${local.slice(6)}`);
  }

  if (digits.length === 10) {
    variants.add(digits);
    variants.add(`+1${digits}`);
    variants.add(`+91${digits}`);
    variants.add(`(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`);
    variants.add(`${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6)}`);
  }

  return [...variants];
}

function buildPhoneFilter(variants) {
  const conditions = variants.map((v) => {
    const escaped = v.replace(/'/g, "''");
    return `(mobilePhone eq '${escaped}' or businessPhones/any(p:p eq '${escaped}'))`;
  });
  return conditions.join(" or ");
}

async function lookupCallerByPhone(accessToken, callerNumber) {
  console.log(`[CALLER-IDENTITY] Looking up caller: ${callerNumber}`);

  const variants = buildPhoneVariants(callerNumber);
  const filter = buildPhoneFilter(variants);
  const url = `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(filter)}&$select=id,displayName&$count=true`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ConsistencyLevel: "eventual",
    },
  });

  const data = await response.json();

  if (data.error) {
    console.error(`[CALLER-IDENTITY] Graph error: ${JSON.stringify(data.error)}`);
    return null;
  }

  if (!data.value || data.value.length === 0) {
    console.log(`[CALLER-IDENTITY] No employee found for: ${callerNumber}`);
    return null;
  }

  return { id: data.value[0].id, displayName: data.value[0].displayName || "Employee" };
}

// Test-only helper: looks a user up by email instead of phone number.
// Never use this for real call authentication - a caller could claim any
// email, whereas the phone number comes from the telecom layer itself.
async function lookupUserByEmail(accessToken, email) {
  const escaped = email.replace(/'/g, "''");
  const url = `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${escaped}'&$select=id,displayName`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  const data = await response.json();

  if (data.error) {
    console.error(`[CALLER-IDENTITY] Graph error looking up by email: ${JSON.stringify(data.error)}`);
    return null;
  }

  if (!data.value || data.value.length === 0) return null;

  return { id: data.value[0].id, displayName: data.value[0].displayName || "Employee" };
}

async function getUserMpin(accessToken, userId) {
  const url = `https://graph.microsoft.com/v1.0/users/${userId}?$select=customSecurityAttributes`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  const data = await response.json();

  if (data.error) {
    console.error(`[CALLER-IDENTITY] Graph error fetching MPIN: ${JSON.stringify(data.error)}`);
    return null;
  }

  // Graph returns the attribute set key as "MPIN" (its actual stored casing),
  // even though the attribute name within it is "Mpin".
  const value = data.customSecurityAttributes?.MPIN?.Mpin;
  return typeof value === "number" ? value : null;
}

// Resolves the caller for a verify_pin request. Real phone calls always use
// the actual caller number. VAPI's dashboard web-test calls never attach a
// customer number at all, so there's nothing to look up by phone in that
// case - as a WEB-TEST-MODE-ONLY fallback, if no number is present and
// ENABLE_WEB_TEST_MODE is explicitly turned on, fall back to a fixed test
// identity (TEST_CALLER_EMAIL) so the PIN flow can still be exercised from
// the dashboard. This never overrides a real phone number, and is a no-op
// unless both env vars are set.
async function resolveCaller(accessToken, callerNumber) {
  if (callerNumber) {
    return lookupCallerByPhone(accessToken, callerNumber);
  }

  if (process.env.ENABLE_WEB_TEST_MODE === "true" && process.env.TEST_CALLER_EMAIL) {
    console.warn(
      `[CALLER-IDENTITY] WEB TEST MODE: no phone number on this call, falling back to TEST_CALLER_EMAIL (${process.env.TEST_CALLER_EMAIL}).`,
    );
    return lookupUserByEmail(accessToken, process.env.TEST_CALLER_EMAIL);
  }

  console.log("[CALLER-IDENTITY] No caller phone number on this call, and web test mode is not enabled.");
  return null;
}

module.exports = { lookupCallerByPhone, lookupUserByEmail, getUserMpin, resolveCaller };
