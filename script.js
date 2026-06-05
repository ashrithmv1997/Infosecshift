// ==========================================
// INFOSEC SHIFT COMMAND CENTER V2 (FIXED)
// ==========================================

// ==========================================
// DOM
// ==========================================

const clock = document.getElementById("clock");
const todayDate = document.getElementById("todayDate");

const activeShiftDisplay = document.getElementById("activeShiftDisplay");
const shiftCountdown = document.getElementById("shiftCountdown");

const s1Count = document.getElementById("s1Count");
const s2Count = document.getElementById("s2Count");
const s3Count = document.getElementById("s3Count");
const offCount = document.getElementById("offCount");

const staffContainer = document.getElementById("staffContainer");

const nextShiftName = document.getElementById("nextShiftName");
const nextShiftTime = document.getElementById("nextShiftTime");
const nextShiftPeople = document.getElementById("nextShiftPeople");

const template = document.getElementById("staffTemplate");

// MODALS
const calendarBtn = document.getElementById("calendarBtn");
const offFinderBtn = document.getElementById("offFinderBtn");

const calendarModal = document.getElementById("calendarModal");
const offModal = document.getElementById("offModal");

const calendarGrid = document.getElementById("calendarGrid");
const selectedDayRoster = document.getElementById("selectedDayRoster");

const employeeSelect = document.getElementById("employeeSelect");
const offResults = document.getElementById("offResults");

// SHIFT ASSISTANT
const assistantBtn = document.getElementById("shiftAssistantBtn");
const assistantPanel = document.getElementById("shiftAssistantPanel");
const assistantContent = document.getElementById("assistantContent");

// NOTICE
const noticeBoard = document.getElementById("shiftNoticeBoard");

// ==========================================
// SHIFT DEFINITIONS
// ==========================================

const SHIFTS = {
  S1: { start: "07:30", end: "15:30", label: "07:30 AM → 03:30 PM" },
  S2: { start: "13:30", end: "22:00", label: "01:30 PM → 10:00 PM" },
  S3: { start: "21:30", end: "08:00", label: "09:30 PM → 08:00 AM" }
};

// ==========================================
// DATA STORE
// ==========================================

let monthRoster = {};
let todayRoster = {};
let employeeNames = [];
let activeCards = [];
let currentDateKey = "";

// ==========================================
// CLOCK
// ==========================================

function updateClock() {
  const now = new Date();

  clock.textContent = now.toLocaleTimeString("en-GB");

  todayDate.textContent = now.toLocaleDateString("en-GB", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric"
  });
}

setInterval(updateClock, 1000);
updateClock();

// ==========================================
// TIME HELPERS
// ==========================================

function excelToDate(serial) {
  return new Date((serial - 25569) * 86400 * 1000);
}

function getDateKey(date) {
  return date.toISOString().split("T")[0];
}

// ⭐ FIX: SHIFT-AWARE DATE (IMPORTANT FIX)
function getShiftAwareDateKey(date) {
  const hours = date.getHours();

  if (hours < 8) {
    const d = new Date(date);
    d.setDate(d.getDate() - 1);
    return getDateKey(d);
  }

  return getDateKey(date);
}

function timeToDate(timeString) {
  const [h, m] = timeString.split(":");
  const d = new Date();
  d.setHours(Number(h), Number(m), 0, 0);
  return d;
}

function formatRemaining(ms) {
  if (ms < 0) ms = 0;

  const total = Math.floor(ms / 1000);
  const hrs = Math.floor(total / 3600);
  const mins = Math.floor((total % 3600) / 60);
  const secs = total % 60;

  return `${String(hrs).padStart(2, "0")}:${String(mins).padStart(2, "0")}:${String(secs).padStart(2, "0")}`;
}

// ==========================================
// ACTIVE SHIFTS
// ==========================================

function getActiveShifts() {
  const now = new Date();
  const mins = now.getHours() * 60 + now.getMinutes();

  const active = [];

  if (mins >= 450 && mins < 930) active.push("S1");
  if (mins >= 810 && mins < 1320) active.push("S2");
  if (mins >= 1290 || mins < 480) active.push("S3");

  return active;
}

// ==========================================
// NEXT BOUNDARY
// ==========================================

function getNextBoundary() {
  const now = new Date();
  const times = ["07:30", "13:30", "15:30", "21:30", "22:00"];

  let nearest = null;

  times.forEach(t => {
    const d = timeToDate(t);

    if (d < now) d.setDate(d.getDate() + 1);

    if (!nearest || d < nearest) nearest = d;
  });

  return nearest;
}

// ==========================================
// HEADER
// ==========================================

function updateHeader() {
  const shifts = getActiveShifts();
  activeShiftDisplay.textContent = shifts.join(" + ");

  const next = getNextBoundary();
  shiftCountdown.textContent = formatRemaining(next - new Date());
}

setInterval(updateHeader, 1000);

// ==========================================
// LOAD EXCEL
// ==========================================

async function loadExcel() {
  const res = await fetch("shifts.xlsx");
  const buffer = await res.arrayBuffer();

  const wb = XLSX.read(buffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  processMonthRoster(rows);
}

// ==========================================
// PROCESS ROSTER
// ==========================================

function processMonthRoster(rows) {
  monthRoster = {};

  const headers = rows[1];

  employeeNames = headers.slice(1, 6);

  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];
    if (!row || typeof row[0] !== "number") continue;

    const date = excelToDate(row[0]);
    const key = getDateKey(date);

    monthRoster[key] = {};

    employeeNames.forEach((emp, idx) => {
      monthRoster[key][emp] = String(row[idx + 1] || "").trim().toUpperCase();
    });
  }

  const todayKey = getShiftAwareDateKey(new Date());
  currentDateKey = todayKey;
  todayRoster = monthRoster[todayKey] || {};

  populateEmployeeDropdown();
}

// ==========================================
// DROPDOWN
// ==========================================

function populateEmployeeDropdown() {
  employeeSelect.innerHTML = `<option value="">Select Employee</option>`;

  employeeNames.forEach(emp => {
    employeeSelect.innerHTML += `<option value="${emp}">${emp}</option>`;
  });
}

// ==========================================
// COUNTS
// ==========================================

function updateCounts() {
  let s1 = 0, s2 = 0, s3 = 0, off = 0;

  Object.values(todayRoster || {}).forEach(sh => {
    if (sh === "S1") s1++;
    else if (sh === "S2") s2++;
    else if (sh === "S3") s3++;
    else off++;
  });

  s1Count.textContent = s1;
  s2Count.textContent = s2;
  s3Count.textContent = s3;
  offCount.textContent = off;
}

// ==========================================
// SAFE TODAY ROSTER USAGE FIX
// ==========================================

function getEmployeesCurrentlyWorking() {
  const active = getActiveShifts();
  const result = [];

  Object.entries(todayRoster || {}).forEach(([name, shift]) => {
    if (active.includes(shift)) {
      result.push({ name, shift });
    }
  });

  return result;
}

// ==========================================
// DASHBOARD
// ==========================================

function buildDashboard() {
  updateCounts();

  staffContainer.innerHTML = "";
  activeCards = [];

  const staff = getEmployeesCurrentlyWorking();

  staff.forEach(s => createStaffCard(s.name, s.shift));

  buildUpcomingShift();
  updateCards();
}

// ==========================================
// STAFF CARD
// ==========================================

function createStaffCard(name, shift) {
  const clone = template.content.cloneNode(true);

  const avatar = clone.querySelector(".avatar");
  const nameEl = clone.querySelector(".staff-name");
  const shiftEl = clone.querySelector(".staff-shift");

  nameEl.textContent = name;

  shiftEl.innerHTML = `${shift}<br><small>${SHIFTS[shift].label}</small>`;

  const key = name.trim().toUpperCase();

  avatar.src =
    PHOTO_MAP?.[key] ||
    `https://ui-avatars.com/api/?name=${encodeURIComponent(name)}`;

  const card = clone.querySelector(".staff-card");

  staffContainer.appendChild(clone);

  activeCards.push({ card, shift });
}

// ==========================================
// CARD PROGRESS
// ==========================================

function updateCards() {
  activeCards.forEach(item => {
    const start = timeToDate(SHIFTS[item.shift].start);
    const end = timeToDate(SHIFTS[item.shift].end);

    const now = new Date();

    const total = end - start;
    const elapsed = now - start;

    let percent = (elapsed / total) * 100;
    percent = Math.max(0, Math.min(100, percent));

    const fill = item.card.querySelector(".progress-fill");
    const text = item.card.querySelector(".progress-text");
    const remain = item.card.querySelector(".remaining-time");

    if (fill) fill.style.width = percent + "%";
    if (text) text.textContent = `${Math.floor(percent)}% Complete`;
    if (remain) remain.textContent = formatRemaining(end - now);
  });
}

setInterval(updateCards, 1000);

// ==========================================
// UPCOMING SHIFT
// ==========================================

function buildUpcomingShift() {
  const now = new Date();
  const mins = now.getHours() * 60 + now.getMinutes();

  let nextShift = "S1";
  let dateKey = currentDateKey;

  if (mins < 810) nextShift = "S2";
  else if (mins < 1320) nextShift = "S3";
  else {
    nextShift = "S1";
    const t = new Date();
    t.setDate(t.getDate() + 1);
    dateKey = getDateKey(t);
  }

  nextShiftName.textContent = nextShift;
  nextShiftTime.textContent = SHIFTS[nextShift].label;

  const roster = monthRoster[dateKey] || {};

  nextShiftPeople.innerHTML = Object.entries(roster)
    .filter(([_, s]) => s === nextShift)
    .map(([n]) => n)
    .join("<br>") || "No Staff";
}

// ==========================================
// SHIFT CHANGE WATCHER
// ==========================================

setInterval(() => {
  const newKey = getShiftAwareDateKey(new Date());

  if (newKey !== currentDateKey) {
    currentDateKey = newKey;
    todayRoster = monthRoster[newKey] || {};
    buildDashboard();
  }
}, 30000);

// ==========================================
// START SYSTEM (ONLY ONCE)
// ==========================================

async function startSystem() {
  await loadExcel();
  buildDashboard();
  updateHeader();
}

document.addEventListener("DOMContentLoaded", startSystem);

// ==========================================
// NOTICE BOARD
// ==========================================

async function loadNotice() {
  try {
    const res = await fetch("https://infosec-notice-api.ashrithmv.workers.dev/notice");
    if (!res.ok) return;

    const notices = await res.json();

    if (!Array.isArray(notices)) return;

    noticeBoard.innerHTML = notices.length
      ? notices.map(n => `
        <div class="notice-card">
          <div>👤 ${n.sender}</div>
          <div>${n.message}</div>
          <div>🕒 ${new Date(n.timestamp).toLocaleString("en-IN")}</div>
        </div>
      `).join("")
      : "<div class='notice-card'>No active notices</div>";

  } catch (e) {
    console.error(e);
  }
}

loadNotice();
setInterval(loadNotice, 10000);
