// testCli.js
// Manual test harness: drives the same lib/ functions the server uses,
// so you can exercise auth -> directory lookup -> availability -> booking
// by typing answers in the terminal, without going through VAPI/ngrok.

require("dotenv").config();
const readline = require("node:readline/promises");
const { stdin, stdout } = require("node:process");

const { getAccessToken } = require("./lib/auth");
const { resolveDate, resolveTime } = require("./lib/dateResolver");
const { resolveAttendees } = require("./lib/directory");
const { checkAvailability, bookMeeting } = require("./lib/calendar");

const DURATION_MINUTES = 30;
const rl = readline.createInterface({ input: stdin, output: stdout });

async function main() {
  console.log("--- Meeting Schedular manual test ---\n");

  const attendeeInput = await rl.question("Attendee name(s), comma separated (blank for none): ");
  const dateInput = await rl.question("Date (e.g. today / tomorrow / friday / 2026-07-20): ");
  const timeInput = await rl.question("Time (e.g. 3pm): ");
  const title = await rl.question("Meeting title: ");

  console.log("\nResolving date/time...");
  const dateResult = resolveDate(dateInput);
  if (dateResult.error) {
    console.log("FAILED:", dateResult.error);
    return rl.close();
  }
  const timeResult = resolveTime(timeInput);
  if (timeResult.error) {
    console.log("FAILED:", timeResult.error);
    return rl.close();
  }
  console.log(`Resolved to: ${dateResult.date} at ${timeResult.time}`);

  console.log("\nAcquiring Graph access token...");
  const accessToken = await getAccessToken();
  if (!accessToken) {
    console.log("FAILED: could not acquire access token. Check TENANT_ID/CLIENT_ID/CLIENT_SECRET in .env");
    return rl.close();
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
      if (proceed.toLowerCase() !== "y") return rl.close();
    }
  }

  console.log("\nChecking availability...");
  const availability = await checkAvailability(accessToken, dateResult.date, timeResult.time, DURATION_MINUTES);
  if (availability.error) {
    console.log("FAILED:", availability.error);
    return rl.close();
  }

  if (!availability.free) {
    console.log(`BUSY. Existing slots that day: ${availability.busySlots.join(", ") || "(none listed)"}`);
    return rl.close();
  }
  console.log("Time is FREE.");

  const confirm = await rl.question("\nBook this meeting now? (y/n): ");
  if (confirm.toLowerCase() !== "y") {
    console.log("Not booked.");
    return rl.close();
  }

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

  rl.close();
}

main().catch((error) => {
  console.error("Unexpected error:", error);
  rl.close();
});
