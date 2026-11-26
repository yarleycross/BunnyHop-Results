"use strict";

const startForm = document.querySelector("#start");
const raceStatus = document.querySelector("#race-status");
const finishes = document.querySelector("#finishes");
const summary = document.querySelector("#summary");
const addFinish = document.querySelector("#add-finish");
const correction = document.querySelector("#correction");
const editStatus = document.querySelector("#edit-status");
const exportForm = document.querySelector("#export");
const exportStatus = document.querySelector("#export-status");
let rows = [];
let editing = null;
let startError = "";

async function requireSuccess(response) {
  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error || "Request failed (" + response.status + ").");
  }
  return response;
}

async function post(url, body) {
  return requireSuccess(await fetch(url, {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify(body),
  }));
}

function elapsed(seconds) {
  if (seconds == null) return "—";
  if (seconds < 0) return "Clock error";
  return Math.floor(seconds / 60) + ":" + String(seconds % 60).padStart(2, "0");
}

function edit(row) {
  editing = row.sequence;
  document.querySelector("#edit-heading").textContent = "Reconcile sequence " + editing;
  document.querySelector("#edit-bib").value = row.correction?.bib ?? "";
  document.querySelector("#edit-seconds").value = row.correction?.elapsed_seconds ?? "";
  document.querySelector("#edit-void").checked = row.void ?? false;
  document.querySelector("#edit-reason").value = "";
  editStatus.textContent = row.correction ? "Last correction: " + row.correction.reason : "";
  correction.hidden = false;
  document.querySelector("#edit-bib").focus();
}

function renderRows(next) {
  if (JSON.stringify(rows) === JSON.stringify(next)) return;
  rows = next;
  const tableRows = rows.map((row) => {
    const tr = document.createElement("tr");
    const incomplete = row.bib == null || row.elapsed_seconds == null || row.elapsed_seconds < 0;
    const status = row.void ? "Voided" : incomplete ? "Needs review" : row.correction ? "Corrected" : "Matched";
    for (const value of [
      row.sequence, elapsed(row.original_seconds), row.original_bib ?? "—",
      elapsed(row.elapsed_seconds), row.bib ?? "—", status,
    ]) {
      const td = tr.insertCell();
      td.textContent = value;
    }
    if (row.finish_ms != null) tr.cells[1].title = new Date(row.finish_ms).toISOString();
    tr.dataset.status = status;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "secondary";
    button.textContent = "Edit";
    button.setAttribute("aria-label", "Edit sequence " + row.sequence);
    button.addEventListener("click", () => edit(row));
    tr.insertCell().append(button);
    return tr;
  });
  finishes.replaceChildren(...tableRows);
}

async function refresh() {
  try {
    const response = await requireSuccess(await fetch("/api/race"));
    const race = await response.json();
    startForm.hidden = race.start_ms !== null;
    addFinish.disabled = race.start_ms === null;
    raceStatus.textContent = race.start_ms === null
      ? startError || "Waiting for the official start."
      : "Started at " + new Date(race.start_ms).toLocaleString();
    renderRows(race.rows);
    const times = rows.filter((row) => row.finish_ms != null).length;
    const bibs = rows.filter((row) => row.original_bib != null).length;
    summary.textContent = times + " captured times · " + bibs + " captured bibs";
  } catch {
    raceStatus.textContent = "Live connection unavailable. Reconnecting…";
  } finally {
    setTimeout(refresh, 1000);
  }
}
refresh();

startForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const earlier = document.querySelector("#start-time");
  const body = {};
  if (event.submitter?.value !== "now") {
    if (!earlier.value) {
      startError = "Enter the earlier official start date and time.";
      raceStatus.textContent = startError;
      return;
    }
    body.start_ms = new Date(earlier.value).getTime();
  }
  if (!window.confirm("Set the official start for this race?")) return;
  try {
    await post("/api/start", body);
    startError = "";
    startForm.hidden = true;
  } catch (error) {
    startError = error.message;
    raceStatus.textContent = startError;
  }
});

addFinish.addEventListener("click", () => edit({sequence: (rows.at(-1)?.sequence ?? 0) + 1}));
document.querySelector("#cancel-edit").addEventListener("click", () => {
  correction.hidden = true;
  editStatus.textContent = "";
});

correction.addEventListener("submit", async (event) => {
  event.preventDefault();
  const seconds = document.querySelector("#edit-seconds").value;
  const submit = event.submitter;
  submit.disabled = true;
  try {
    await post("/api/correction/" + editing, {
      elapsed_seconds: seconds === "" ? null : Number(seconds),
      bib: document.querySelector("#edit-bib").value || null,
      void: document.querySelector("#edit-void").checked,
      reason: document.querySelector("#edit-reason").value,
    });
    correction.hidden = true;
    editStatus.textContent = "Correction saved.";
  } catch (error) {
    editStatus.textContent = error.message;
  } finally {
    submit.disabled = false;
  }
});

exportForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const submit = event.submitter;
  submit.disabled = true;
  exportStatus.textContent = "";
  try {
    const response = await requireSuccess(await fetch("/export", {
      method: "POST",
      body: new FormData(exportForm),
    }));
    const url = URL.createObjectURL(await response.blob());
    const link = document.createElement("a");
    link.href = url;
    link.download = "Bunny Hop Live Results Entry.xlsx";
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    exportStatus.textContent = "Workbook downloaded.";
  } catch (error) {
    exportStatus.textContent = error.message;
  } finally {
    submit.disabled = false;
  }
});
