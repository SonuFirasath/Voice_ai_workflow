// server.js
// Standalone tool server for ElevenLabs Conversational AI - exposes
// check_availability and book_appointment as separate HTTP endpoints,
// meant to be tunneled with ngrok and registered as two Webhook tools on
// an ElevenLabs agent.
//
// Unlike the VAPI server in ../Meeting Schedular, ElevenLabs webhook tools
// send the declared parameters directly as the request body and read the
// response body directly - there is no toolCallList/results envelope and
// no per-call id to track PIN verification against. This server reuses the
// same Microsoft Graph logic (auth/date-resolving/directory/calendar) from
// ../Meeting Schedular/lib but skips PIN verification entirely - there is
// no verify_pin tool on this agent.

require("dotenv").config();
const express = require("express");

const { getAccessToken } = require("../Meeting Schedular/lib/auth");
const { resolveDate, resolveTime } = require("../Meeting Schedular/lib/dateResolver");
const { resolveAttendees } = require("../Meeting Schedular/lib/directory");
const { checkAvailability, bookMeeting } = require("../Meeting Schedular/lib/calendar");

const app = express();
app.use(express.json());

const DURATION_MINUTES = 30;

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
  res.json({ status: "ElevenLabs Schedular is running." });
});

app.post("/check-availability", async (req, res) => {
  try {
    const { date, time, attendees } = req.body || {};
    console.log(`[EL-SCHEDULER] check_availability args: date="${date}" time="${time}" attendees=${JSON.stringify(attendees)}`);

    const dateResult = resolveDate(date);
    if (dateResult.error) return res.json({ result: dateResult.error });

    const timeResult = resolveTime(time);
    if (timeResult.error) return res.json({ result: timeResult.error });

    const accessToken = await getAccessToken();
    if (!accessToken) return res.json({ result: "Calendar authentication failed. Please try again." });

    const attendeeNames = Array.isArray(attendees) ? attendees : [];
    const { resolved, unresolved } = await resolveAttendees(accessToken, attendeeNames);

    const availability = await checkAvailability(
      accessToken,
      dateResult.date,
      timeResult.time,
      DURATION_MINUTES,
      resolved.map((r) => r.email),
    );
    if (availability.error) return res.json({ result: availability.error });

    const parts = [];
    if (unresolved.length > 0) {
      parts.push(`Could not find a directory match for: ${unresolved.map((u) => u.name).join(", ")}.`);
    }
    parts.push(
      availability.free
        ? `${dateResult.date} at ${timeResult.time} is available for everyone.`
        : `${dateResult.date} at ${timeResult.time} has a conflict: ${describeConflicts(availability.conflicts, resolved)}.`,
    );

    res.json({ result: parts.join(" ") });
  } catch (error) {
    console.error("[EL-SCHEDULER] check_availability failed:", error.message);
    res.json({ result: "Something went wrong checking the calendar. Please try again." });
  }
});

app.post("/book-appointment", async (req, res) => {
  try {
    const { date, time, title, attendees } = req.body || {};
    console.log(`[EL-SCHEDULER] book_appointment args: date="${date}" time="${time}" title="${title}" attendees=${JSON.stringify(attendees)}`);

    const dateResult = resolveDate(date);
    if (dateResult.error) return res.json({ result: dateResult.error });

    const timeResult = resolveTime(time);
    if (timeResult.error) return res.json({ result: timeResult.error });

    if (!title) return res.json({ result: "A meeting title is required before booking." });

    const accessToken = await getAccessToken();
    if (!accessToken) return res.json({ result: "Calendar authentication failed. Please try again." });

    const attendeeNames = Array.isArray(attendees) ? attendees : [];
    const { resolved, unresolved } = await resolveAttendees(accessToken, attendeeNames);

    if (unresolved.length > 0) {
      const names = unresolved.map((u) => u.name).join(", ");
      return res.json({
        result: `Could not find a directory match for: ${names}. Please confirm the correct name or provide their email address before booking.`,
      });
    }

    const availability = await checkAvailability(
      accessToken,
      dateResult.date,
      timeResult.time,
      DURATION_MINUTES,
      resolved.map((r) => r.email),
    );
    if (availability.error) return res.json({ result: availability.error });
    if (!availability.free) {
      return res.json({
        result: `Can't book ${dateResult.date} at ${timeResult.time}: ${describeConflicts(availability.conflicts, resolved)}. Please pick another time.`,
      });
    }

    const booking = await bookMeeting(accessToken, {
      dateStr: dateResult.date,
      timeStr: timeResult.time,
      durationMinutes: DURATION_MINUTES,
      title,
      attendeeEmails: resolved.map((r) => r.email),
    });

    if (booking.error) return res.json({ result: booking.error });

    res.json({
      result: `Meeting "${title}" booked for ${dateResult.date} at ${timeResult.time} with ${resolved.map((r) => r.displayName).join(", ") || "no additional attendees"}.`,
    });
  } catch (error) {
    console.error("[EL-SCHEDULER] book_appointment failed:", error.message);
    res.json({ result: "Something went wrong booking the meeting. Please try again." });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`[EL-SCHEDULER] ElevenLabs Schedular listening on port ${PORT}`);
});
