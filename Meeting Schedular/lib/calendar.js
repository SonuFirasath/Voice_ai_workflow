// lib/calendar.js
// Microsoft Graph calendar operations against the shared target mailbox.
// Dates/times in and out of this module are plain "YYYY-MM-DD" / "HH:mm:ss"
// wall-clock strings in TIME_ZONE - Graph is told the timeZone separately,
// so no UTC conversion is needed here.

const TARGET_EMAIL = process.env.TARGET_EMAIL;
const TIME_ZONE = process.env.TIME_ZONE || "Asia/Kolkata";
const WORKING_HOURS_START = "09:00:00";
const WORKING_HOURS_END = "17:00:00";

function timeStringToMinutes(hhmmss) {
  const [h, m] = hhmmss.split(":").map(Number);
  return h * 60 + m;
}

function extractMinutesFromDateTime(dateTimeStr) {
  // Graph returns "2026-07-20T10:00:00.0000000" (no offset) when a Prefer
  // timeZone header is sent - the HH:mm is already in TIME_ZONE.
  const match = dateTimeStr.match(/T(\d{2}):(\d{2})/);
  return match ? parseInt(match[1], 10) * 60 + parseInt(match[2], 10) : null;
}

function pad(n) {
  return String(n).padStart(2, "0");
}

function addMinutes(dateStr, timeStr, minutesToAdd) {
  const [h, m, s] = timeStr.split(":").map(Number);
  let total = h * 60 + m + minutesToAdd;
  const dayOffset = Math.floor(total / 1440);
  total = ((total % 1440) + 1440) % 1440;

  let newDateStr = dateStr;
  if (dayOffset !== 0) {
    const d = new Date(`${dateStr}T00:00:00`);
    d.setDate(d.getDate() + dayOffset);
    newDateStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  return { dateStr: newDateStr, timeStr: `${pad(Math.floor(total / 60))}:${pad(total % 60)}:${pad(s || 0)}` };
}

async function getScheduleItems(accessToken, dateStr) {
  const endpoint = `https://graph.microsoft.com/v1.0/users/${TARGET_EMAIL}/calendar/getSchedule`;

  const requestBody = {
    schedules: [TARGET_EMAIL],
    startTime: { dateTime: `${dateStr}T${WORKING_HOURS_START}`, timeZone: TIME_ZONE },
    endTime: { dateTime: `${dateStr}T${WORKING_HOURS_END}`, timeZone: TIME_ZONE },
    availabilityViewInterval: 30,
  };

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Prefer: `outlook.timezone="${TIME_ZONE}"`,
    },
    body: JSON.stringify(requestBody),
  });

  const data = await response.json();

  if (data.error) {
    console.error("[CALENDAR] Graph error:", data.error);
    return { error: "Failed to access the calendar schedule." };
  }

  return { scheduleItems: data.value?.[0]?.scheduleItems || [] };
}

async function checkAvailability(accessToken, dateStr, timeStr, durationMinutes = 30) {
  console.log(`[CALENDAR] Checking availability on ${dateStr} at ${timeStr}`);

  const { scheduleItems, error } = await getScheduleItems(accessToken, dateStr);
  if (error) return { error };

  const requestedStart = timeStringToMinutes(timeStr);
  const requestedEnd = requestedStart + durationMinutes;

  const busySlots = scheduleItems.map(
    (item) => `${item.start.dateTime.slice(11, 16)}-${item.end.dateTime.slice(11, 16)}`,
  );

  const conflict = scheduleItems.some((item) => {
    const busyStart = extractMinutesFromDateTime(item.start.dateTime);
    const busyEnd = extractMinutesFromDateTime(item.end.dateTime);
    return requestedStart < busyEnd && busyStart < requestedEnd;
  });

  return { free: !conflict, busySlots };
}

async function bookMeeting(accessToken, { dateStr, timeStr, durationMinutes = 30, title, attendeeEmails = [] }) {
  console.log(`[CALENDAR] Booking "${title}" on ${dateStr} at ${timeStr} with ${attendeeEmails.length} attendee(s)`);

  const { dateStr: endDateStr, timeStr: endTimeStr } = addMinutes(dateStr, timeStr, durationMinutes);

  const requestBody = {
    subject: title,
    start: { dateTime: `${dateStr}T${timeStr}`, timeZone: TIME_ZONE },
    end: { dateTime: `${endDateStr}T${endTimeStr}`, timeZone: TIME_ZONE },
    attendees: attendeeEmails.map((email) => ({
      emailAddress: { address: email },
      type: "required",
    })),
  };

  const response = await fetch(`https://graph.microsoft.com/v1.0/users/${TARGET_EMAIL}/events`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(requestBody),
  });

  const data = await response.json();

  if (data.error) {
    console.error("[CALENDAR] Graph error:", data.error);
    return { error: "Failed to create the calendar event." };
  }

  return { eventId: data.id, webLink: data.webLink };
}

module.exports = { checkAvailability, bookMeeting };
