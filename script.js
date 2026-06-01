// ==========================================
// SHIFT COMMAND CENTER
// script.js
// ==========================================

// ----------------------------
// DOM
// ----------------------------

const clock = document.getElementById("clock");
const todayDate = document.getElementById("todayDate");
const currentShiftEl = document.getElementById("currentShift");
const shiftCountdown = document.getElementById("shiftCountdown");

const s1Count = document.getElementById("s1Count");
const s2Count = document.getElementById("s2Count");
const s3Count = document.getElementById("s3Count");
const offCount = document.getElementById("offCount");

const staffContainer = document.getElementById("staffContainer");

const nextShiftName = document.getElementById("nextShiftName");
const nextShiftTime = document.getElementById("nextShiftTime");
const nextShiftPeople = document.getElementById("nextShiftPeople");

const excelFile = document.getElementById("excelFile");

const radarS1 = document.getElementById("radarS1");
const radarS2 = document.getElementById("radarS2");
const radarS3 = document.getElementById("radarS3");
const radarOFF = document.getElementById("radarOFF");

const template = document.getElementById("staffTemplate");

// ----------------------------
// SHIFT DEFINITIONS
// ----------------------------

const SHIFTS = {
    S1: {
        start: "07:30",
        end: "15:30"
    },
    S2: {
        start: "13:30",
        end: "22:00"
    },
    S3: {
        start: "21:30",
        end: "08:00"
    }
};
const SHIFT_DISPLAY = {

    S1: "07:30 AM → 03:30 PM",

    S2: "01:30 PM → 10:00 PM",

    S3: "09:30 PM → 08:00 AM"

};

// ----------------------------
// GLOBALS
// ----------------------------

let todayRoster = {};
let activeCards = [];

// ----------------------------
// LIVE CLOCK
// ----------------------------

function updateClock() {

    const now = new Date();

    clock.textContent =
        now.toLocaleTimeString("en-GB");

    todayDate.textContent =
        now.toLocaleDateString("en-GB", {
            weekday: "long",
            day: "numeric",
            month: "long",
            year: "numeric"
        });

}

setInterval(updateClock, 1000);
updateClock();

// ----------------------------
// TIME HELPERS
// ----------------------------

function timeToDate(timeStr) {

    const [h, m] = timeStr.split(":");

    const d = new Date();

    d.setHours(Number(h));
    d.setMinutes(Number(m));
    d.setSeconds(0);

    return d;
}

function formatRemaining(ms) {

    if (ms < 0) ms = 0;

    const totalSec = Math.floor(ms / 1000);

    const hrs = Math.floor(totalSec / 3600);
    const mins = Math.floor((totalSec % 3600) / 60);
    const secs = totalSec % 60;

    return (
        String(hrs).padStart(2, "0") +
        ":" +
        String(mins).padStart(2, "0") +
        ":" +
        String(secs).padStart(2, "0")
    );
}

// ----------------------------
// CURRENT SHIFT
// ----------------------------

function getCurrentShift() {

    const now = new Date();

    const currentMinutes =
        now.getHours() * 60 +
        now.getMinutes();

    if (
        currentMinutes >= 450 &&
        currentMinutes < 930
    ) {
        return "S1";
    }

    if (
        currentMinutes >= 810 &&
        currentMinutes < 1320
    ) {
        return "S2";
    }

    return "S3";
}

// ----------------------------
// SHIFT END
// ----------------------------

function getShiftEnd(shift) {

    const now = new Date();

    let end;

    if (shift === "S1") {

        end = timeToDate("15:30");

    } else if (shift === "S2") {

        end = timeToDate("22:00");

    } else {

        end = timeToDate("08:00");

        if (
            now.getHours() >= 21
        ) {
            end.setDate(
                end.getDate() + 1
            );
        }
    }

    return end;
}

// ----------------------------
// SHIFT HEADER
// ----------------------------

function updateShiftHeader() {

    const shift =
        getCurrentShift();

    currentShiftEl.textContent =
        shift;

    radarS1.classList.remove("active");
    radarS2.classList.remove("active");
    radarS3.classList.remove("active");

    document
        .getElementById("radar" + shift)
        .classList.add("active");

    const end =
        getShiftEnd(shift);

    const remaining =
        end - new Date();

    shiftCountdown.textContent =
        formatRemaining(
            remaining
        );
}

setInterval(
    updateShiftHeader,
    1000
);

updateShiftHeader();

// ----------------------------
// EXCEL IMPORT
// ----------------------------

excelFile.addEventListener(
    "change",
    loadExcel
);

function loadExcel(event) {

    const file =
        event.target.files[0];

    if (!file) return;

    const reader =
        new FileReader();

    reader.onload =
        function (e) {

            const data =
                new Uint8Array(
                    e.target.result
                );

            const workbook =
                XLSX.read(data, {
                    type: "array"
                });

            const sheet =
                workbook.Sheets[
                    workbook.SheetNames[0]
                ];

            const rows =
                XLSX.utils.sheet_to_json(
                    sheet,
                    {
                        header: 1
                    }
                );

            processRoster(rows);
        };

    reader.readAsArrayBuffer(
        file
    );
}

// ----------------------------
// PROCESS EXCEL
// ----------------------------

function processRoster(rows) {

    todayRoster = {};

    // Employee names are always row 2 in your sheet
    const headers = rows[1];

    const employeeNames = [
        headers[1],
        headers[2],
        headers[3],
        headers[4],
        headers[5]
    ];

    const now = new Date();

    const todayDay = now.getDate();

    let todayRow = null;

    // Search actual date rows
    for (let i = 2; i < rows.length; i++) {

        const cell = rows[i][0];

        if (!cell) continue;

        let dayNumber = null;

        if (cell instanceof Date) {
            dayNumber = cell.getDate();
        }

        else if (!isNaN(Date.parse(cell))) {
            dayNumber = new Date(cell).getDate();
        }

        if (dayNumber === todayDay) {
            todayRow = rows[i];
            break;
        }
    }

    if (!todayRow) {

        alert("Today's roster not found");

        return;
    }

    employeeNames.forEach((name, index) => {

        const shift = todayRow[index + 1];

        todayRoster[name] = shift;

    });

    buildDashboard();
}

// ----------------------------
// COUNTS
// ----------------------------

function updateCounts() {

    let s1 = 0;
    let s2 = 0;
    let s3 = 0;
    let off = 0;

    Object.values(
        todayRoster
    ).forEach(shift => {

        if (shift === "S1")
            s1++;

        else if (
            shift === "S2"
        )
            s2++;

        else if (
            shift === "S3"
        )
            s3++;

        else
            off++;
    });

    s1Count.textContent = s1;
    s2Count.textContent = s2;
    s3Count.textContent = s3;
    offCount.textContent = off;
}

// ----------------------------
// STAFF CARDS
// ----------------------------

function buildDashboard() {

    updateCounts();
updateOffPeople();
    const currentShift =
        getCurrentShift();

    staffContainer.innerHTML =
        "";

    activeCards = [];

    Object.entries(
        todayRoster
    ).forEach(
        ([name, shift]) => {

            if (
                shift !==
                currentShift
            )
                return;

            const clone =
                template.content.cloneNode(
                    true
                );

            clone.querySelector(
                ".staff-name"
            ).textContent = name;

            clone.querySelector(".staff-shift").innerHTML = `
${shift}
<br>
<small>${SHIFT_DISPLAY[shift]}</small>
`;

            clone.querySelector(
                ".avatar"
            ).src =
                `https://ui-avatars.com/api/?background=00d9ff&color=fff&name=${encodeURIComponent(
                    name
                )}`;

            const card =
                clone.querySelector(
                    ".staff-card"
                );

            staffContainer.appendChild(
                clone
            );

            activeCards.push({
                card,
                shift
            });
        }
    );

    buildNextShift();

    updateCards();
}

// ----------------------------
// CARD LIVE UPDATE
// ----------------------------

function updateCards() {

    const currentShift =
        getCurrentShift();

    const shiftInfo =
        SHIFTS[
            currentShift
        ];

    const start =
        timeToDate(
            shiftInfo.start
        );

    const end =
        getShiftEnd(
            currentShift
        );

    activeCards.forEach(
        item => {

            const progressFill =
                item.card.querySelector(
                    ".progress-fill"
                );

            const progressText =
                item.card.querySelector(
                    ".progress-text"
                );

            const remainingText =
                item.card.querySelector(
                    ".remaining-time"
                );

            const total =
                end - start;

            const done =
                new Date() -
                start;

            let percent =
                (
                    done /
                    total
                ) *
                100;

            percent =
                Math.max(
                    0,
                    Math.min(
                        100,
                        percent
                    )
                );

            progressFill.style.width =
                percent + "%";

            progressText.textContent =
                Math.floor(
                    percent
                ) + "%";

            remainingText.textContent =
                formatRemaining(
                    end -
                        new Date()
                ) +
                " Remaining";
        }
    );
}

setInterval(
    updateCards,
    1000
);

// ----------------------------
// NEXT SHIFT
// ----------------------------

function buildNextShift() {

    const current =
        getCurrentShift();

    let next;

    if (current === "S1")
        next = "S2";
    else if (
        current === "S2"
    )
        next = "S3";
    else next = "S1";

    nextShiftName.textContent =
        next;

    nextShiftTime.textContent =
        SHIFTS[next].start;

    const people =
        Object.entries(
            todayRoster
        )
            .filter(
                ([n, s]) =>
                    s === next
            )
            .map(
                ([n]) => n
            );

    nextShiftPeople.innerHTML =
        people.length
            ? people.join(
                  "<br>"
              )
            : "No staff";
}

// ----------------------------
// DEMO DATA
// ----------------------------

todayRoster = {

    "Ashwin": "S1",
    "Rahul R G": "S2",
    "Hari": "S3",
    "Ashrith": "S2",
    "Manu": "OFF"

};

buildDashboard();

// ----------------------------
// AUTO REFRESH SHIFT CHANGE
// ----------------------------

let lastShift =
    getCurrentShift();

setInterval(() => {

    const current =
        getCurrentShift();

    if (
        current !==
        lastShift
    ) {

        lastShift =
            current;

        buildDashboard();
    }

}, 30000);
function updateOffPeople() {

    const container =
        document.getElementById("offPeople");

    container.innerHTML = "";

    Object.entries(todayRoster)
        .forEach(([name, shift]) => {

            if (shift === "OFF") {

                const div =
                    document.createElement("div");

                div.className =
                    "off-person";

                div.textContent =
                    name;

                container.appendChild(div);
            }
        });
}