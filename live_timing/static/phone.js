"use strict";

const form = document.querySelector("#capture");
const stream = form.dataset.stream;
const bib = document.querySelector("#bib");
const button = document.querySelector("#record");
const notice = document.querySelector("#notice");
const last = document.querySelector("#last");
const key = "bunnyhop-pending-" + stream;
const label = stream === "finish" ? "FINISH" : "Save bib";
let pending = JSON.parse(localStorage.getItem(key));
let saving = false;

function updateButton() {
  button.textContent = saving ? "Saving…" : pending ? "Retry save" : label;
  button.disabled = saving;
  if (bib) bib.readOnly = Boolean(pending);
}

if (pending) {
  if (bib) bib.value = pending.bib;
  notice.textContent = "A save is unconfirmed. Retry it before recording another runner.";
}
updateButton();

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (saving) return;
  if (!pending) {
    // getRandomValues works over private Wi-Fi HTTP.
    const bytes = crypto.getRandomValues(new Uint8Array(16));
    const requestId = Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
    pending = {request_id: requestId, ...(bib ? {bib: bib.value} : {})};
    localStorage.setItem(key, JSON.stringify(pending));
  }
  saving = true;
  updateButton();
  notice.textContent = "";
  try {
    const response = await fetch("/api/capture/" + stream, {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(pending),
      signal: AbortSignal.timeout(8000),
    });
    const result = await response.json();
    if (!response.ok) {
      if (response.status < 500 && !result.retryable) {
        localStorage.removeItem(key);
        pending = null;
      }
      notice.textContent = result.error;
      return;
    }
    localStorage.removeItem(key);
    pending = null;
    last.textContent = stream === "finish"
      ? "Saved finish #" + result.sequence
      : "Saved bib " + result.bib + " at #" + result.sequence;
    if (bib) bib.value = "";
  } catch {
    notice.textContent = "Save unconfirmed. Retry this entry before recording another runner.";
  } finally {
    saving = false;
    updateButton();
    if (bib && !pending) bib.focus();
  }
});
