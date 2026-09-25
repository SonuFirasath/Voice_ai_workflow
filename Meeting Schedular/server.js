// server.js
// Standalone tool server for VAPI - exposes check_availability and
// book_appointment as separate HTTP endpoints, meant to be tunneled with
// ngrok and registered as two Custom Tools on the VAPI assistant.

require("dotenv").config();
const express = require("express");

const { getAccessToken } = require("./lib/auth");
const { resolveDate, resolveTime } = require("./lib/dateResolver");
const { resolveAttendees } = require("./lib/directory");
const { checkAvailability, bookMeeting } = require("./lib/calendar");
const { resolveCaller, getUserMpin } = require("./lib/callerIdentity");
const { isVerified, isLockedOut, markVerified, recordFailedAttempt } = require("./lib/callState");

const app = express();
app.use(express.json());

const DURATION_MINUTES = 30;

function extractToolCalls(body) {
  return body?.message?.toolCallList || body?.message?.toolCalls || [];
}

function parseArgs(toolCall) {
  const raw = toolCall.function?.arguments ?? toolCall.arguments;
  if (typeof raw === "string") {
    try {
      return JSON.parse(raw);
    } catch {
      return {};
    }
  }
  return raw || {};
}

// Turns checkAvailability's conflict list into a caller-friendly sentence,
// naming the specific person (or "the calendar" for the organizer) who's busy.
function describeConflicts(conflicts, resolvedAttendees) {
  return conflicts
    .map((c) => {
      if (c.email.toLowerCase() === (process.env.TARGET_EMAIL || "").toLowerCase()) {
        return `the calendar is busy at ${c.busySlots.join(", ")}`;
      }
      const match = resolvedAttendees.find((r) => r.email.toLowerCase() === c.email.toLowerCase());
      const label = match ? match.displayName : c.email;
      return `${label} is busy at ${c.busySlots.join(", ")}`;
    })
    .join("; ");
}

app.get("/", (req, res) => {
  res.json({ status: "Meeting Schedular is running." });
});

app.post("/verify-pin", async (req, res) => {
  const toolCalls = extractToolCalls(req.body);
  const callId = req.body?.message?.call?.id;
  const callerNumber = req.body?.message?.call?.customer?.number;
  const results = [];

  for (const toolCall of toolCalls) {
    try {
      const rawArgs = parseArgs(toolCall);
      const { pin } = rawArgs;
      console.log(`[SCHEDULER] verify_pin raw toolCall: ${JSON.stringify(toolCall)}`);
      console.log(`[SCHEDULER] verify_pin parsed args: ${JSON.stringify(rawArgs)} callId="${callId}"`);

      if (!callId) {
        console.warn("[SCHEDULER] verify_pin: no callId on request, cannot track attempts.");
        results.push({ toolCallId: toolCall.id, result: "Could not identify this call. Please try again." });
        continue;
      }

      if (isLockedOut(callId)) {
        console.warn(`[SCHEDULER] verify_pin: callId ${callId} is already locked out.`);
        results.push({
          toolCallId: toolCall.id,
          result: 'This caller already failed PIN verification twice. End the call now and tell them: "Your PIN is incorrect, please try again some other time."',
        });
        continue;
      }

      const digits = String(pin ?? "").replace(/\D/g, "");
      if (!digits) {
        console.warn(`[SCHEDULER] verify_pin: no digits extracted from pin="${pin}"`);
        results.push({ toolCallId: toolCall.id, result: "No PIN digits were received. Ask the caller to enter their 4-digit PIN again." });
        continue;
      }

      const accessToken = await getAccessToken();
      if (!accessToken) {
        console.warn("[SCHEDULER] verify_pin: getAccessToken() failed.");
        results.push({ toolCallId: toolCall.id, result: "Authentication service unavailable. Please try again." });
        continue;
      }

      const caller = await resolveCaller(accessToken, callerNumber);
      if (!caller) {
        console.warn(`[SCHEDULER] verify_pin: no employee matched (callerNumber=${callerNumber || "none"})`);
        results.push({
          toolCallId: toolCall.id,
          result: "This caller could not be matched to an employee record. End the call and tell them their number isn't recognized.",
        });
        continue;
      }

      const storedPin = await getUserMpin(accessToken, caller.id);
      if (storedPin === null) {
        console.warn(`[SCHEDULER] verify_pin: no MPIN configured for ${caller.displayName} (${caller.id})`);
        results.push({
          toolCallId: toolCall.id,
          result: "No PIN is configured for this account. End the call and ask them to contact an administrator.",
        });
        continue;
      }

      if (parseInt(digits, 10) === storedPin) {
        markVerified(callId);
        console.log(`[SCHEDULER] verify_pin: SUCCESS for ${caller.displayName} (callId ${callId})`);
        results.push({
          toolCallId: toolCall.id,
          result: `PIN verified for ${caller.displayName}. The caller is authenticated and may now proceed with scheduling.`,
        });
        continue;
      }

      const state = recordFailedAttempt(callId);
      console.warn(`[SCHEDULER] verify_pin: WRONG PIN for ${caller.displayName} (attempt ${state.attempts}, lockedOut=${state.lockedOut})`);
      results.push({
        toolCallId: toolCall.id,
        result: state.lockedOut
          ? 'Incorrect PIN, second attempt failed. End the call now and tell the caller: "Your PIN is incorrect, please try again some other time."'
          : "That PIN is incorrect. Ask the caller to enter their 4-digit PIN one more time.",
      });
    } catch (error) {
      console.error("[SCHEDULER] verify_pin failed:", error.message);
      results.push({ toolCallId: toolCall.id, result: "Something went wrong verifying the PIN. Please try again." });
    }
  }

  res.json({ results });
});

app.post("/check-availability", async (req, res) => {
  const toolCalls = extractToolCalls(req.body);
  const callId = req.body?.message?.call?.id;
  const results = [];

  for (const toolCall of toolCalls) {
    try {
      if (!isVerified(callId)) {
        results.push({
          toolCallId: toolCall.id,
          result: "The caller has not completed PIN verification yet. Verify their PIN before checking the calendar.",
        });
        continue;
      }

      const { date, time, attendees } = parseArgs(toolCall);
      console.log(`[SCHEDULER] check_availability args: date="${date}" time="${time}" attendees=${JSON.stringify(attendees)}`);

      const dateResult = resolveDate(date);
      if (dateResult.error) {
        results.push({ toolCallId: toolCall.id, result: dateResult.error });
        continue;
      }

      const timeResult = resolveTime(time);
      if (timeResult.error) {
        results.push({ toolCallId: toolCall.id, result: timeResult.error });
        continue;
      }

      const accessToken = await getAccessToken();
      if (!accessToken) {
        results.push({ toolCallId: toolCall.id, result: "Calendar authentication failed. Please try again." });
        continue;
      }

      const attendeeNames = Array.isArray(attendees) ? attendees : [];
      const { resolved, unresolved } = await resolveAttendees(accessToken, attendeeNames);

      const availability = await checkAvailability(
        accessToken,
        dateResult.date,
        timeResult.time,
        DURATION_MINUTES,
        resolved.map((r) => r.email),
      );
      if (availability.error) {
        results.push({ toolCallId: toolCall.id, result: availability.error });
        continue;
      }

      const parts = [];
      if (unresolved.length > 0) {
        parts.push(`Could not find a directory match for: ${unresolved.map((u) => u.name).join(", ")}.`);
      }
      parts.push(
        availability.free
          ? `${dateResult.date} at ${timeResult.time} is available for everyone.`
          : `${dateResult.date} at ${timeResult.time} has a conflict: ${describeConflicts(availability.conflicts, resolved)}.`,
      );

      results.push({ toolCallId: toolCall.id, result: parts.join(" ") });
    } catch (error) {
      console.error("[SCHEDULER] check_availability failed:", error.message);
      results.push({ toolCallId: toolCall.id, result: "Something went wrong checking the calendar. Please try again." });
    }
  }

  res.json({ results });
});

app.post("/book-appointment", async (req, res) => {
  const toolCalls = extractToolCalls(req.body);
  const callId = req.body?.message?.call?.id;
  const results = [];

  for (const toolCall of toolCalls) {
    try {
      if (!isVerified(callId)) {
        results.push({
          toolCallId: toolCall.id,
          result: "The caller has not completed PIN verification yet. Verify their PIN before booking anything.",
        });
        continue;
      }

      const { date, time, title, attendees } = parseArgs(toolCall);
      console.log(`[SCHEDULER] book_appointment args: date="${date}" time="${time}" title="${title}" attendees=${JSON.stringify(attendees)}`);

      const dateResult = resolveDate(date);
      if (dateResult.error) {
        results.push({ toolCallId: toolCall.id, result: dateResult.error });
        continue;
      }

      const timeResult = resolveTime(time);
      if (timeResult.error) {
        results.push({ toolCallId: toolCall.id, result: timeResult.error });
        continue;
      }

      if (!title) {
        results.push({ toolCallId: toolCall.id, result: "A meeting title is required before booking." });
        continue;
      }

      const accessToken = await getAccessToken();
      if (!accessToken) {
        results.push({ toolCallId: toolCall.id, result: "Calendar authentication failed. Please try again." });
        continue;
      }

      const attendeeNames = Array.isArray(attendees) ? attendees : [];
      const { resolved, unresolved } = await resolveAttendees(accessToken, attendeeNames);

      if (unresolved.length > 0) {
        const names = unresolved.map((u) => u.name).join(", ");
        results.push({
          toolCallId: toolCall.id,
          result: `Could not find a directory match for: ${names}. Please confirm the correct name or provide their email address before booking.`,
        });
        continue;
      }

      const availability = await checkAvailability(
        accessToken,
        dateResult.date,
        timeResult.time,
        DURATION_MINUTES,
        resolved.map((r) => r.email),
      );
      if (availability.error) {
        results.push({ toolCallId: toolCall.id, result: availability.error });
        continue;
      }
      if (!availability.free) {
        results.push({
          toolCallId: toolCall.id,
          result: `Can't book ${dateResult.date} at ${timeResult.time}: ${describeConflicts(availability.conflicts, resolved)}. Please pick another time.`,
        });
        continue;
      }

      const booking = await bookMeeting(accessToken, {
        dateStr: dateResult.date,
        timeStr: timeResult.time,
        durationMinutes: DURATION_MINUTES,
        title,
        attendeeEmails: resolved.map((r) => r.email),
      });

      if (booking.error) {
        results.push({ toolCallId: toolCall.id, result: booking.error });
        continue;
      }

      results.push({
        toolCallId: toolCall.id,
        result: `Meeting "${title}" booked for ${dateResult.date} at ${timeResult.time} with ${resolved.map((r) => r.displayName).join(", ") || "no additional attendees"}.`,
      });
    } catch (error) {
      console.error("[SCHEDULER] book_appointment failed:", error.message);
      results.push({ toolCallId: toolCall.id, result: "Something went wrong booking the meeting. Please try again." });
    }
  }

  res.json({ results });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`[SCHEDULER] Meeting Schedular listening on port ${PORT}`);
});
