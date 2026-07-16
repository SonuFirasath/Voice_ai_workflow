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

app.get("/", (req, res) => {
  res.json({ status: "Meeting Schedular is running." });
});

app.post("/check-availability", async (req, res) => {
  const toolCalls = extractToolCalls(req.body);
  const results = [];

  for (const toolCall of toolCalls) {
    try {
      const { date, time } = parseArgs(toolCall);
      console.log(`[SCHEDULER] check_availability args: date="${date}" time="${time}"`);

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

      const availability = await checkAvailability(accessToken, dateResult.date, timeResult.time, DURATION_MINUTES);
      if (availability.error) {
        results.push({ toolCallId: toolCall.id, result: availability.error });
        continue;
      }

      const message = availability.free
        ? `${dateResult.date} at ${timeResult.time} is available.`
        : `${dateResult.date} at ${timeResult.time} is already booked. Busy slots that day: ${availability.busySlots.join(", ") || "none listed"}.`;

      results.push({ toolCallId: toolCall.id, result: message });
    } catch (error) {
      console.error("[SCHEDULER] check_availability failed:", error.message);
      results.push({ toolCallId: toolCall.id, result: "Something went wrong checking the calendar. Please try again." });
    }
  }

  res.json({ results });
});

app.post("/book-appointment", async (req, res) => {
  const toolCalls = extractToolCalls(req.body);
  const results = [];

  for (const toolCall of toolCalls) {
    try {
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

      const availability = await checkAvailability(accessToken, dateResult.date, timeResult.time, DURATION_MINUTES);
      if (availability.error) {
        results.push({ toolCallId: toolCall.id, result: availability.error });
        continue;
      }
      if (!availability.free) {
        results.push({
          toolCallId: toolCall.id,
          result: `${dateResult.date} at ${timeResult.time} is already booked. Busy slots that day: ${availability.busySlots.join(", ") || "none listed"}. Please pick another time.`,
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
