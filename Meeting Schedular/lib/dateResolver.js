// lib/dateResolver.js
// Turns the relative phrases a caller naturally says ("tomorrow", "next Friday")
// into a concrete date, and a spoken time ("3pm") into an HH:mm:ss string.
// Only resolves single, unambiguous days - ranges like "this week"/"this month"
// are expected to have already been narrowed to one day by the assistant.

const WEEKDAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];

function formatDateLocal(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function resolveDate(phraseRaw, referenceDate = new Date()) {
  const phrase = (phraseRaw || "").trim().toLowerCase();

  if (!phrase) return { error: "No date provided." };

  if (phrase === "today") {
    return { date: formatDateLocal(referenceDate) };
  }

  if (phrase === "tomorrow") {
    const d = new Date(referenceDate);
    d.setDate(d.getDate() + 1);
    return { date: formatDateLocal(d) };
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(phrase)) {
    return { date: phrase };
  }

  const weekdayMatch = phrase.match(/^(next\s+|this\s+)?(sunday|monday|tuesday|wednesday|thursday|friday|saturday)$/);
  if (weekdayMatch) {
    const wantsNextWeek = weekdayMatch[1]?.trim() === "next";
    const targetIdx = WEEKDAYS.indexOf(weekdayMatch[2]);
    const todayIdx = referenceDate.getDay();

    let diff = (targetIdx - todayIdx + 7) % 7;
    // Someone naming a weekday almost always means the *next* occurrence of it,
    // not "today" - so treat diff === 0 as "same day next week" unless resolved elsewhere.
    if (diff === 0) diff = 7;
    if (wantsNextWeek) diff += 7;

    const d = new Date(referenceDate);
    d.setDate(d.getDate() + diff);
    return { date: formatDateLocal(d) };
  }

  // Fallback: explicit dates like "July 20" or "7/20/2026".
  if (/[a-z]{3,}\s+\d{1,2}/.test(phrase) || /\d{1,2}\/\d{1,2}(\/\d{2,4})?/.test(phrase)) {
    const parsed = new Date(phrase);
    if (!Number.isNaN(parsed.getTime())) {
      return { date: formatDateLocal(parsed) };
    }
  }

  return { error: `Could not understand the date "${phraseRaw}". Please ask the caller for a specific day.` };
}

function resolveTime(phraseRaw) {
  const phrase = (phraseRaw || "").trim().toLowerCase();
  const match = phrase.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/);

  if (!match) {
    return { error: `Could not understand the time "${phraseRaw}". Please ask the caller for a specific time.` };
  }

  let hour = parseInt(match[1], 10);
  const minute = match[2] ? parseInt(match[2], 10) : 0;
  const ampm = match[3];

  if (ampm === "am") {
    if (hour === 12) hour = 0;
  } else if (ampm === "pm") {
    if (hour !== 12) hour += 12;
  }

  if (hour > 23 || minute > 59) {
    return { error: `Could not understand the time "${phraseRaw}". Please ask the caller for a specific time.` };
  }

  return { time: `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}:00` };
}

module.exports = { resolveDate, resolveTime };
