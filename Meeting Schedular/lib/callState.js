// lib/callState.js
// In-memory, per-call PIN verification state, keyed by VAPI's call id.
// Not persisted - fine for this use case since a call's entire lifetime
// (PIN entry through booking) happens within one running server process.

const MAX_ATTEMPTS = 2;
const TTL_MS = 30 * 60 * 1000; // stale entries (abandoned calls) expire after 30 min

const calls = new Map();

function sweepExpired() {
  const now = Date.now();
  for (const [callId, state] of calls) {
    if (now - state.updatedAt > TTL_MS) calls.delete(callId);
  }
}

function isVerified(callId) {
  sweepExpired();
  return !!calls.get(callId)?.verified;
}

function isLockedOut(callId) {
  sweepExpired();
  return !!calls.get(callId)?.lockedOut;
}

function markVerified(callId) {
  calls.set(callId, { attempts: 0, verified: true, lockedOut: false, updatedAt: Date.now() });
}

function recordFailedAttempt(callId) {
  const state = calls.get(callId) || { attempts: 0, verified: false, lockedOut: false };
  state.attempts += 1;
  state.updatedAt = Date.now();
  if (state.attempts >= MAX_ATTEMPTS) state.lockedOut = true;
  calls.set(callId, state);
  return state;
}

module.exports = { isVerified, isLockedOut, markVerified, recordFailedAttempt, MAX_ATTEMPTS };
