// testCli.js
// Manual test harness: drives the same lib/ functions the server uses,
// so you can exercise auth -> directory lookup -> availability -> booking,
// or PIN verification, by typing answers in the terminal, without going
// through VAPI/ngrok.

require("dotenv").config();
const readline = require("node:readline/promises");
const { stdin, stdout } = require("node:process");

const { getAccessToken } = require("./lib/auth");
const { resolveDate, resolveTime } = require("./lib/dateResolver");
const { resolveAttendees } = require("./lib/directory");
const { checkAvailability, bookMeeting } = require("./lib/calendar");
const { lookupUserByEmail, getUserMpin } = require("./lib/callerIdentity");

const DURATION_MINUTES = 30;
const rl = readline.createInterface({ input: stdin, output: stdout });

async function testMeetingScheduling() {
  const attendeeInput = await rl.question("Attendee name(s), comma separated (blank for none): ");
  const dateInput = await rl.question("Date (e.g. today / tomorrow / friday / 2026-07-20): ");
  const timeInput = await rl.question("Time (e.g. 3pm): ");
  const title = await rl.question("Meeting title: ");

  console.log("\nResolving date/time...");
  const dateResult = resolveDate(dateInput);
  if (dateResult.error) return console.log("FAILED:", dateResult.error);

  const timeResult = resolveTime(timeInput);
  if (timeResult.error) return console.log("FAILED:", timeResult.error);
  console.log(`Resolved to: ${dateResult.date} at ${timeResult.time}`);

  console.log("\nAcquiring Graph access token...");
  const accessToken = await getAccessToken();
  if (!accessToken) {
    return console.log("FAILED: could not acquire access token. Check TENANT_ID/CLIENT_ID/CLIENT_SECRET in .env");
  }
  console.log("Token acquired.");

  const attendeeNames = attendeeInput
    .split(",")
    .map((n) => n.trim())
    .filter(Boolean);

  let resolvedAttendees = [];
  if (attendeeNames.length > 0) {
    console.log("\nResolving attendees against the directory...");
    const { resolved, unresolved } = await resolveAttendees(accessToken, attendeeNames);
    resolvedAttendees = resolved;

    for (const r of resolved) console.log(`  MATCHED  "${r.name}" -> ${r.displayName} <${r.email}>`);
    for (const u of unresolved) console.log(`  NOT FOUND "${u.name}": ${u.error}`);

    if (unresolved.length > 0) {
      const proceed = await rl.question("\nSome attendees didn't resolve. Continue anyway? (y/n): ");
      if (proceed.toLowerCase() !== "y") return;
    }
  }

  console.log("\nChecking availability (organizer + attendees)...");
  const availability = await checkAvailability(
    accessToken,
    dateResult.date,
    timeResult.time,
    DURATION_MINUTES,
    resolvedAttendees.map((r) => r.email),
  );
  if (availability.error) return console.log("FAILED:", availability.error);

  if (!availability.free) {
    for (const c of availability.conflicts) {
      const match = resolvedAttendees.find((r) => r.email.toLowerCase() === c.email.toLowerCase());
      console.log(`BUSY: ${match ? match.displayName : c.email} at ${c.busySlots.join(", ")}`);
    }
    return;
  }
  console.log("Time is FREE for everyone.");

  const confirm = await rl.question("\nBook this meeting now? (y/n): ");
  if (confirm.toLowerCase() !== "y") return console.log("Not booked.");

  console.log("\nBooking...");
  const booking = await bookMeeting(accessToken, {
    dateStr: dateResult.date,
    timeStr: timeResult.time,
    durationMinutes: DURATION_MINUTES,
    title,
    attendeeEmails: resolvedAttendees.map((r) => r.email),
  });

  if (booking.error) {
    console.log("FAILED:", booking.error);
  } else {
    console.log(`BOOKED. Event ID: ${booking.eventId}`);
    console.log(`Link: ${booking.webLink}`);
  }
}

async function testPinVerification() {
  const emailInput = await rl.question("Email to test (blank for FirasathG@chimeratechnologies.com): ");
  const email = emailInput.trim() || "FirasathG@chimeratechnologies.com";
  const pinInput = await rl.question("PIN to test: ");

  console.log("\nAcquiring Graph access token...");
  const accessToken = await getAccessToken();
  if (!accessToken) {
    return console.log("FAILED: could not acquire access token. Check TENANT_ID/CLIENT_ID/CLIENT_SECRET in .env");
  }
  console.log("Token acquired.");

  console.log(`\nLooking up ${email}...`);
  const user = await lookupUserByEmail(accessToken, email);
  if (!user) return console.log(`FAILED: no user found for ${email}`);
  console.log(`Found: ${user.displayName} (${user.id})`);

  console.log("\nFetching MPIN custom security attribute...");
  const storedPin = await getUserMpin(accessToken, user.id);
  if (storedPin === null) return console.log("FAILED: no MPIN configured for this account.");
  console.log(`Stored MPIN: ${storedPin}`);

  const digits = pinInput.replace(/\D/g, "");
  if (!digits) return console.log("No digits entered.");

  const isMatch = parseInt(digits, 10) === storedPin;
  console.log(isMatch ? "\nMATCH - PIN is correct." : "\nNO MATCH - PIN is incorrect.");
}

async function main() {
  console.log("--- Meeting Schedular manual test ---\n");
  const mode = await rl.question("Test (1) PIN verification or (2) meeting scheduling? [1/2]: ");

  if (mode.trim() === "1") {
    await testPinVerification();
  } else {
    await testMeetingScheduling();
  }

  rl.close();
}

main().catch((error) => {
  console.error("Unexpected error:", error);
  rl.close();
});