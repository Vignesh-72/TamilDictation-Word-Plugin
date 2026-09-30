/*
 * commands.js — Ribbon command function file for Tamil Dictation Word Add-in
 *
 * FIX #10: Removed the boilerplate Outlook Mailbox API usage
 * (Office.context.mailbox) which is undefined in Word add-ins and would throw
 * a runtime error if this function was ever invoked from the ribbon.
 *
 * This file is intentionally minimal. All dictation logic lives in the
 * taskpane (App.jsx). The ribbon button's sole action is ShowTaskpane,
 * so no custom command function is currently needed.
 */

/* global Office */

Office.onReady(() => {
  // Office.js is ready. No initialization needed for this add-in.
});

// Required boilerplate: associate the function with Office actions.
// If a ribbon FunctionFile command is ever added, implement it here using
// Word-appropriate APIs (e.g., Word.run), NOT Outlook Mailbox APIs.
function action(event) {
  // Currently unused — the ribbon button uses ShowTaskpane, not ExecuteFunction.
  // Safely complete the event to avoid Office hanging.
  event.completed();
}

// Register the function so Office can locate it if needed.
Office.actions.associate("action", action);
