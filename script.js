// ==========================================
// INFOSEC SHIFT COMMAND CENTER V2
// script.js PART 1
// CORE ENGINE + EXCEL PARSER
// ==========================================

// ==========================================
// DOM
// ==========================================

const clock =
document.getElementById("clock");

const todayDate =
document.getElementById("todayDate");

const activeShiftDisplay =
document.getElementById(
    "activeShiftDisplay"
);

const shiftCountdown =
document.getElementById(
    "shiftCountdown"
);

const s1Count =
document.getElementById(
    "s1Count"
);

const s2Count =
document.getElementById(
    "s2Count"
);

const s3Count =
document.getElementById(
    "s3Count"
);

const offCount =
document.getElementById(
    "offCount"
);

const staffContainer =
document.getElementById(
    "staffContainer"
);

const nextShiftName =
document.getElementById(
    "nextShiftName"
);

const nextShiftTime =
document.getElementById(
    "nextShiftTime"
);

const nextShiftPeople =
document.getElementById(
    "nextShiftPeople"
);

const template =
document.getElementById(
    "staffTemplate"
);

// ==========================================
// MODALS
// ==========================================

const calendarBtn =
document.getElementById(
    "calendarBtn"
);

const offFinderBtn =
document.getElementById(
    "offFinderBtn"
);

const calendarModal =
document.getElementById(
    "calendarModal"
);

const offModal =
document.getElementById(
    "offModal"
);

const calendarGrid =
document.getElementById(
    "calendarGrid"
);

const selectedDayRoster =
document.getElementById(
    "selectedDayRoster"
);

const employeeSelect =
document.getElementById(
    "employeeSelect"
);

const offResults =
document.getElementById(
    "offResults"
);

// ==========================================
// SHIFT DEFINITIONS
// ==========================================

const SHIFTS = {

    S1: {

        start: "07:30",

        end: "15:30",

        label:
        "07:30 AM → 03:30 PM"
    },

    S2: {

        start: "13:30",

        end: "22:00",

        label:
        "01:30 PM → 10:00 PM"
    },

    S3: {

        start: "21:30",

        end: "08:00",

        label:
        "09:30 PM → 08:00 AM"
    }
};

// ==========================================
// EMPLOYEE PHOTOS
// ==========================================

const PHOTO_MAP = {

    "ASHWIN":
    "ashwin.jpg",

    "RAHUL R G":
    "rahul.jpg",

    "HARI":
    "hari.jpg",

    "ASHRITH":
    "ashrith.jpg",

    "MANU":
    "manu.jpg"
};

// ==========================================
// GLOBAL DATA STORE
// ==========================================

let monthRoster = {};

let todayRoster = {};

let employeeNames = [];

let activeCards = [];

let currentDateKey = "";

// ==========================================
// CLOCK
// ==========================================

function updateClock(){

    const now =
    new Date();

    clock.textContent =
    now.toLocaleTimeString(
        "en-GB"
    );

    todayDate.textContent =
    now.toLocaleDateString(
        "en-GB",
        {
            weekday:"long",
            day:"numeric",
            month:"long",
            year:"numeric"
        }
    );
}

setInterval(
    updateClock,
    1000
);

updateClock();

// ==========================================
// EXCEL DATE HELPER
// ==========================================

function excelToDate(
    serial
){

    return new Date(

        (
            serial - 25569
        )
        *
        86400
        *
        1000
    );
}

// ==========================================
// DATE KEY
// YYYY-MM-DD
// ==========================================

function getDateKey(
    date
){

    const y =
    date.getFullYear();

    const m =
    String(
        date.getMonth()+1
    )
    .padStart(2,"0");

    const d =
    String(
        date.getDate()
    )
    .padStart(2,"0");

    return `${y}-${m}-${d}`;
}

// ==========================================
// TIME HELPERS
// ==========================================

function timeToDate(
    timeString
){

    const [h,m] =
    timeString.split(":");

    const d =
    new Date();

    d.setHours(
        Number(h)
    );

    d.setMinutes(
        Number(m)
    );

    d.setSeconds(0);

    d.setMilliseconds(0);

    return d;
}

function formatRemaining(
    ms
){

    if(ms < 0)
        ms = 0;

    const total =
    Math.floor(ms/1000);

    const hrs =
    Math.floor(
        total / 3600
    );

    const mins =
    Math.floor(
        (
            total % 3600
        ) / 60
    );

    const secs =
    total % 60;

    return (

        String(hrs)
        .padStart(2,"0")

        + ":"

        +

        String(mins)
        .padStart(2,"0")

        + ":"

        +

        String(secs)
        .padStart(2,"0")
    );
}

// ==========================================
// ACTIVE SHIFTS
// SUPPORTS OVERLAP
// ==========================================

function getActiveShifts(){

    const now =
    new Date();

    const mins =

        now.getHours()
        *
        60

        +

        now.getMinutes();

    const active = [];

    // S1
    if(
        mins >= 450 &&
        mins < 930
    ){

        active.push(
            "S1"
        );
    }

    // S2
    if(
        mins >= 810 &&
        mins < 1320
    ){

        active.push(
            "S2"
        );
    }

    // S3
    if(

        mins >= 1290

        ||

        mins < 480

    ){

        active.push(
            "S3"
        );
    }

    return active;
}

// ==========================================
// NEXT SHIFT CHANGE
// ==========================================

function getNextBoundary(){

    const now =
    new Date();

    const candidates = [

        "07:30",

        "13:30",

        "15:30",

        "21:30",

        "22:00"
    ];

    let nearest =
    null;

    candidates.forEach(
        time=>{

            const d =
            timeToDate(
                time
            );

            if(
                d < now
            ){

                d.setDate(
                    d.getDate()+1
                );
            }

            if(

                !nearest

                ||

                d < nearest

            ){

                nearest = d;
            }
        }
    );

    return nearest;
}

// ==========================================
// HEADER UPDATE
// ==========================================

function updateHeader(){

    const shifts =
    getActiveShifts();

    activeShiftDisplay
    .textContent =

    shifts.join(
        " + "
    );

    const next =
    getNextBoundary();

    shiftCountdown
    .textContent =

    formatRemaining(
        next - new Date()
    );
}

setInterval(
    updateHeader,
    1000
);

updateHeader();

// ==========================================
// AUTO LOAD EXCEL
// shifts.xlsx
// ==========================================

async function loadExcel(){

    try{

        const response =
        await fetch(
            "shifts.xlsx"
        );

        const buffer =
        await response
        .arrayBuffer();

        const workbook =
        XLSX.read(
            buffer,
            {
                type:"array"
            }
        );

        const sheet =
        workbook.Sheets[
            workbook.SheetNames[0]
        ];

        const rows =
        XLSX.utils.sheet_to_json(
            sheet,
            {
                header:1
            }
        );

        processMonthRoster(
            rows
        );
    }

    catch(error){

        console.error(
            error
        );

        alert(
            "Unable to load shifts.xlsx"
        );
    }
}

// ==========================================
// BUILD ENTIRE MONTH
// FROM EXCEL
// ==========================================

function processMonthRoster(
    rows
){

    monthRoster = {};

    const headers =
    rows[1];

    employeeNames = [

        headers[1],

        headers[2],

        headers[3],

        headers[4],

        headers[5]
    ];

    for(

        let i = 2;

        i < rows.length;

        i++

    ){

        const row =
        rows[i];

        if(
            !row
            ||
            typeof row[0]
            !== "number"
        ){

            continue;
        }

        const date =
        excelToDate(
            row[0]
        );

        const rosterDate =
new Date();

if(
    rosterDate.getHours() < 8
){

    rosterDate.setDate(
        rosterDate.getDate() - 1
    );
}

const todayKey =
getDateKey(
    rosterDate
);

        monthRoster[
            key
        ] = {};

        employeeNames.forEach(

            (
                employee,
                index
            )=>{

                monthRoster[
                    key
                ][
                    employee
                ] =

                String(
                    row[
                        index+1
                    ]
                )

                .trim()

                .toUpperCase();
            }
        );
    }

    const todayKey =
    getDateKey(
        new Date()
    );

    currentDateKey =
    todayKey;

    todayRoster =
    monthRoster[
        todayKey
    ] || {};
console.log(
    "Today Roster:",
    JSON.stringify(todayRoster, null, 2)
);
    console.log(
        "Month Roster Loaded",
        monthRoster
    );

    console.log(
        "Today's Roster",
        todayRoster
    );

    populateEmployeeDropdown();
}

// ==========================================
// POPULATE OFF FINDER
// ==========================================

function populateEmployeeDropdown(){

    employeeSelect.innerHTML =

    `
    <option value="">
    Select Employee
    </option>
    `;

    employeeNames.forEach(
        employee=>{

            employeeSelect
            .innerHTML +=

            `
            <option value="${employee}">
            ${employee}
            </option>
            `;
        }
    );
}

// ==========================================
// PART 1 END
// ==========================================
//
// NEXT:
// Generate V2 script.js Part 2
//
// Part 2 includes:
//
// updateCounts()
// buildDashboard()
// buildStaffCards()
// next OFF calculation
// upcoming shift card
// progress bars
// auto shift refresh
//
// ==========================================
// INFOSEC SHIFT COMMAND CENTER V2
// script.js PART 2
// DASHBOARD + STAFF CARDS
// ==========================================

// ==========================================
// SHIFT END TIME
// ==========================================

function getShiftEnd(shift){

    const now = new Date();

    let end;

    if(shift === "S1"){

        end = timeToDate("15:30");
    }

    else if(shift === "S2"){

        end = timeToDate("22:00");
    }

    else{

        end = timeToDate("08:00");

        // If shift started tonight,
        // end is tomorrow morning

        if(now.getHours() >= 21){

            end.setDate(
                end.getDate() + 1
            );
        }
    }

    return end;
}

// ==========================================
// SHIFT START TIME
// ==========================================

function getShiftStart(shift){

    const now = new Date();

    let start;

    if(shift === "S1"){

        start = timeToDate("07:30");
    }

    else if(shift === "S2"){

        start = timeToDate("13:30");
    }

    else{

        start = timeToDate("21:30");

        // After midnight but before 8 AM
        if(now.getHours() < 8){

            start.setDate(
                start.getDate() - 1
            );
        }
    }

    return start;
}

// ==========================================
// COUNTERS
// ==========================================

function updateCounts(){

    let s1 = 0;
    let s2 = 0;
    let s3 = 0;
    let off = 0;

    Object.values(todayRoster)
    .forEach(shift=>{

        if(shift === "S1")
            s1++;

        else if(shift === "S2")
            s2++;

        else if(shift === "S3")
            s3++;

        else
            off++;
    });

    s1Count.textContent = s1;
    s2Count.textContent = s2;
    s3Count.textContent = s3;
    offCount.textContent = off;
}

// ==========================================
// OFF TODAY SECTION
// ==========================================

function updateOffPeople(){

    const container =
    document.getElementById(
        "offPeople"
    );

    if(!container) return;

    container.innerHTML = "";

    Object.entries(todayRoster)
    .forEach(([name,shift])=>{

        if(shift === "OFF"){

            const div =
            document.createElement(
                "div"
            );

            div.className =
            "off-person";

            div.textContent =
            name;

            container.appendChild(
                div
            );
        }
    });
}

// ==========================================
// NEXT OFF CALCULATOR
// ==========================================

function findNextOff(employee){

    const dates =
    Object.keys(monthRoster)
    .sort();

    const today =
    currentDateKey;

    let found = null;

    for(let i=0;i<dates.length;i++){

        const date = dates[i];

        if(date < today)
            continue;

        const shift =
        monthRoster[date][employee];

        if(shift === "OFF"){

            found = date;
            break;
        }
    }

    if(!found)
        return "No OFF";

    const d =
    new Date(found);

    return d.toLocaleDateString(
        "en-GB",
        {
            day:"numeric",
            month:"short",
            weekday:"short"
        }
    );
}

// ==========================================
// ACTIVE STAFF
// ==========================================

function getEmployeesCurrentlyWorking(){

    const activeShifts =
    getActiveShifts();

    const result = [];

    Object.entries(todayRoster)
    .forEach(([name,shift])=>{

        if(
            activeShifts.includes(
                shift
            )
        ){

            result.push({

                name,
                shift
            });
        }
    });

    return result;
}

// ==========================================
// BUILD DASHBOARD
// ==========================================

function buildDashboard(){

    updateCounts();

    updateOffPeople();

    staffContainer.innerHTML =
    "";

    activeCards = [];

    const staff =
    getEmployeesCurrentlyWorking();

    staff.forEach(employee=>{

        createStaffCard(
            employee.name,
            employee.shift
        );
    });

    buildUpcomingShift();

    updateCards();
}

// ==========================================
// CREATE STAFF CARD
// ==========================================

function createStaffCard(
    name,
    shift
){

    const clone =
    template.content.cloneNode(
        true
    );

    const avatar =
    clone.querySelector(
        ".avatar"
    );

    const nameEl =
    clone.querySelector(
        ".staff-name"
    );

    const shiftEl =
    clone.querySelector(
        ".staff-shift"
    );

    const nextOffEl =
    clone.querySelector(
        ".next-off"
    );

    nameEl.textContent =
    name;

    shiftEl.innerHTML =

    `
    ${shift}
    <br>
    <small>
    ${SHIFTS[shift].label}
    </small>
    `;

    const key =
    name.toUpperCase();

    avatar.src =

    PHOTO_MAP[key]

    ||

    `https://ui-avatars.com/api/?background=00d9ff&color=fff&name=${encodeURIComponent(name)}`;

    if(nextOffEl){

        nextOffEl.innerHTML =

        `
        Next OFF:
        <br>
        ${findNextOff(name)}
        `;
    }

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

// ==========================================
// PROGRESS ENGINE
// ==========================================

function updateCards(){

    activeCards.forEach(item=>{

        const shift =
        item.shift;

        const start =
        getShiftStart(
            shift
        );

        const end =
        getShiftEnd(
            shift
        );

        const now =
        new Date();

        const total =
        end - start;

        const elapsed =
        now - start;

        let percent =

        (
            elapsed
            /
            total
        )

        * 100;

        percent =
        Math.max(
            0,
            Math.min(
                100,
                percent
            )
        );

        const fill =
        item.card.querySelector(
            ".progress-fill"
        );

        const text =
        item.card.querySelector(
            ".progress-text"
        );

        const remain =
        item.card.querySelector(
            ".remaining-time"
        );

        if(fill){

            fill.style.width =
            percent + "%";
        }

        if(text){

            text.textContent =

            Math.floor(percent)

            + "% Complete";
        }

        if(remain){

            remain.textContent =

            formatRemaining(
                end - now
            )

            +

            " Remaining";
        }
    });
}

// ==========================================
// LIVE CARD REFRESH
// ==========================================

setInterval(
    updateCards,
    1000
);

// ==========================================
// UPCOMING SHIFT
// ==========================================

function getNextShift(){

    const active =
    getActiveShifts();

    if(
        active.includes("S1")
        &&
        active.includes("S2")
    ){

        return "S3";
    }

    if(
        active.includes("S2")
        &&
        active.includes("S3")
    ){

        return "S1";
    }

    if(active.includes("S1"))
        return "S2";

    if(active.includes("S2"))
        return "S3";

    return "S1";
}

// ==========================================
// NEXT SHIFT CARD
// ==========================================

function buildUpcomingShift(){

    const now = new Date();

    let nextShift;
    let rosterDateKey;

    const mins =
        now.getHours() * 60 +
        now.getMinutes();

    // S1 active
    if(
        mins >= 450 &&
        mins < 810
    ){

        nextShift = "S2";
        rosterDateKey = currentDateKey;
    }

    // S1 + S2 overlap
    else if(
        mins >= 810 &&
        mins < 930
    ){

        nextShift = "S3";
        rosterDateKey = currentDateKey;
    }

    // S2 only
    else if(
        mins >= 930 &&
        mins < 1290
    ){

        nextShift = "S3";
        rosterDateKey = currentDateKey;
    }

    // S2 + S3 overlap
    else if(
        mins >= 1290 &&
        mins < 1320
    ){

        nextShift = "S1";

        const tomorrow =
            new Date();

        tomorrow.setDate(
            tomorrow.getDate() + 1
        );

        rosterDateKey =
            getDateKey(tomorrow);
    }

    // S3 only
    else{

        nextShift = "S1";

        const tomorrow =
            new Date();

        if(mins >= 1320){

            tomorrow.setDate(
                tomorrow.getDate() + 1
            );
        }

        rosterDateKey =
            getDateKey(tomorrow);
    }

    nextShiftName.textContent =
        nextShift;

    nextShiftTime.textContent =
        SHIFTS[nextShift].label;

    const roster =
        monthRoster[
            rosterDateKey
        ] || {};

    const people =
        Object.entries(roster)
        .filter(
            ([name,shift]) =>
            shift === nextShift
        )
        .map(
            ([name]) => name
        );

    nextShiftPeople.innerHTML =

        people.length

        ?

        people.join("<br>")

        :

        "No Staff";
}
// ==========================================
// SHIFT CHANGE WATCHER
// ==========================================

let lastShiftState =

JSON.stringify(
    getActiveShifts()
);

setInterval(()=>{

    const current =

    JSON.stringify(
        getActiveShifts()
    );

    if(
        current !==
        lastShiftState
    ){

        lastShiftState =
        current;

        buildDashboard();
    }

},15000);

// ==========================================
// INITIAL DASHBOARD
// ==========================================

function initializeDashboard(){

    buildDashboard();

    updateCards();

    buildUpcomingShift();
}

// ==========================================
// PART 2 END
// ==========================================
//
// NEXT:
//
// Generate V2 script.js Part 3
//
// Includes:
//
// Calendar modal
// Date click roster
// OFF Finder
// Modal controls
// Startup sequence
// loadExcel()
// initializeDashboard()
//
// ==========================================
// INFOSEC SHIFT COMMAND CENTER V2
// script.js PART 3
// CALENDAR + OFF FINDER + STARTUP
// ==========================================

// ==========================================
// CALENDAR MODAL OPEN
// ==========================================

if(calendarBtn){

    calendarBtn.addEventListener(
        "click",
        ()=>{

            buildCalendar();

            calendarModal.classList.add(
                "show"
            );
        }
    );
}

// ==========================================
// OFF FINDER MODAL OPEN
// ==========================================

if(offFinderBtn){

    offFinderBtn.addEventListener(
        "click",
        ()=>{

            offModal.classList.add(
                "show"
            );
        }
    );
}

// ==========================================
// CLOSE BUTTONS
// ==========================================

document
.querySelectorAll(".close-btn")
.forEach(btn=>{

    btn.addEventListener(
        "click",
        ()=>{

            calendarModal.classList.remove(
                "show"
            );

            offModal.classList.remove(
                "show"
            );
        }
    );
});

// ==========================================
// CLICK OUTSIDE MODAL
// ==========================================

window.addEventListener(
    "click",
    e=>{

        if(
            e.target === calendarModal
        ){

            calendarModal.classList.remove(
                "show"
            );
        }

        if(
            e.target === offModal
        ){

            offModal.classList.remove(
                "show"
            );
        }
    }
);

// ==========================================
// BUILD MONTH CALENDAR
// ==========================================

function buildCalendar(){

    if(!calendarGrid)
        return;

    calendarGrid.innerHTML = "";

    const dates =
    Object.keys(monthRoster)
    .sort();

    dates.forEach(dateKey=>{

        const date =
        new Date(dateKey);

        const day =
        date.getDate();

        const weekday =
        date.toLocaleDateString(
            "en-GB",
            {
                weekday:"short"
            }
        );

        const div =
        document.createElement(
            "div"
        );

        div.className =
        "calendar-day";

        div.innerHTML =

        `
        <strong>${day}</strong>
        <br>
        <small>${weekday}</small>
        `;

        div.addEventListener(
            "click",
            ()=>{

                showRosterForDate(
                    dateKey
                );
            }
        );

        calendarGrid.appendChild(
            div
        );
    });
}

// ==========================================
// SHOW ROSTER OF SELECTED DAY
// ==========================================

function showRosterForDate(
    dateKey
){

    const roster =
    monthRoster[
        dateKey
    ];

    if(
        !roster
    ){

        selectedDayRoster.innerHTML =
        "No roster found";

        return;
    }

    let s1 = [];
    let s2 = [];
    let s3 = [];
    let off = [];

    Object.entries(roster)
    .forEach(
        ([name,shift])=>{

            if(
                shift === "S1"
            ){

                s1.push(name);
            }

            else if(
                shift === "S2"
            ){

                s2.push(name);
            }

            else if(
                shift === "S3"
            ){

                s3.push(name);
            }

            else{

                off.push(name);
            }
        }
    );

    selectedDayRoster.innerHTML =

    `
    <h3>
    ${new Date(dateKey)
        .toLocaleDateString(
            "en-GB",
            {
                weekday:"long",
                day:"numeric",
                month:"long"
            }
        )}
    </h3>

    <br>

    <b>S1</b><br>
    ${s1.join(", ") || "-"}

    <br><br>

    <b>S2</b><br>
    ${s2.join(", ") || "-"}

    <br><br>

    <b>S3</b><br>
    ${s3.join(", ") || "-"}

    <br><br>

    <b>OFF</b><br>
    ${off.join(", ") || "-"}
    `;
}

// ==========================================
// OFF FINDER
// ==========================================

if(employeeSelect){

    employeeSelect.addEventListener(
        "change",
        ()=>{
            showEmployeeOffs();
        }
    );
}

// ==========================================
// SHOW ALL OFFS
// ==========================================

function showEmployeeOffs(){

    const employee =
    employeeSelect.value;

    if(
        !employee
    ){

        offResults.innerHTML =

        "Select an employee";

        return;
    }

    const offDays = [];

    Object.keys(monthRoster)
    .sort()
    .forEach(dateKey=>{

        const shift =

        monthRoster[
            dateKey
        ][
            employee
        ];

        if(
            shift === "OFF"
        ){

            const date =
            new Date(
                dateKey
            );

            offDays.push(

                `
                <div class="off-person">
                ${date.toLocaleDateString(
                    "en-GB",
                    {
                        weekday:"short",
                        day:"numeric",
                        month:"short"
                    }
                )}
                </div>
                `
            );
        }
    });

    offResults.innerHTML =

    `
    <h3>
    ${employee}
    </h3>

    <br>

    ${

        offDays.length

        ?

        offDays.join("")

        :

        "No OFF days found"

    }
    `;
}

// ==========================================
// DAILY ROSTER REFRESH
// MIDNIGHT AUTO UPDATE
// ==========================================

function checkDateChange(){

    const now = new Date();

    let rosterDate = new Date(now);

    // S3 continues until 08:00 AM
    if(now.getHours() < 8){

        rosterDate.setDate(
            rosterDate.getDate() - 1
        );
    }

    const newKey =
    getDateKey(rosterDate);

    if(
        newKey !== currentDateKey
    ){

        currentDateKey =
        newKey;

        todayRoster =

        monthRoster[
            newKey
        ] || {};

        buildDashboard();
    }
}
// ==========================================
// LOAD DATA
// ==========================================

async function startSystem(){

    await loadExcel();

    initializeDashboard();

    console.log(
        "SHIFT COMMAND CENTER V2 READY"
    );
}

startSystem();

// ==========================================
// OPTIONAL AUTO RELOAD
// EVERY 5 MINUTES
// Useful when shifts.xlsx
// is replaced on server
// ==========================================

setInterval(
    async ()=>{

        try{

            await loadExcel();

            buildDashboard();

            console.log(
                "Roster refreshed"
            );
        }

        catch(err){

            console.error(
                err
            );
        }

    },

    300000
);
// ==========================================
// LOAD DATA
// ==========================================

async function startSystem(){

    await loadExcel();

    initializeDashboard();

    console.log(
        "SHIFT COMMAND CENTER V2 READY"
    );
}

startSystem();


// ==========================================
// AUTO REFRESH WHEN SHIFT CHANGES
// ==========================================

function checkShiftBoundary() {

    const now = new Date();

    const hh = now.getHours();
    const mm = now.getMinutes();

    if (
    (hh === 7  && mm >= 30 && mm <= 31) ||
    (hh === 13 && mm >= 30 && mm <= 31) ||
    (hh === 21 && mm >= 30 && mm <= 31)
) {

        console.log(
            "Shift boundary reached"
        );

        loadExcel();

        buildDashboard();
    }
}

setInterval(
    checkShiftBoundary,
    30000
);


// ==========================================
// REFRESH WHEN APP/TAB BECOMES ACTIVE
// ==========================================

document.addEventListener(
    "visibilitychange",
    () => {

        if (!document.hidden) {

            console.log(
                "App resumed"
            );

            loadExcel();

            buildDashboard();
        }
    }
);
// ==========================================
// END OF V2 SCRIPT
// ==========================================
//
// FINAL FOLDER STRUCTURE:
//
// /index.html
// /style.css
// /script.js
// /shifts.xlsx
//
// /photos
//    ashwin.jpg
//    rahul.jpg
//    hari.jpg
//    ashrith.jpg
//    manu.jpg
//
// Upload all to GitHub.
// Connect GitHub to Cloudflare Pages.
// Replace shifts.xlsx every month.
// Site updates automatically.
// ==========================================
// =====================================
// SHIFT ASSISTANT
// =====================================

const assistantBtn =
document.getElementById(
    "shiftAssistantBtn"
);

const assistantPanel =
document.getElementById(
    "shiftAssistantPanel"
);

const assistantContent =
document.getElementById(
    "assistantContent"
);

assistantBtn.addEventListener(
    "click",
    ()=>{

        if(
            assistantPanel.style.display
            === "block"
        ){

            assistantPanel.style.display =
            "none";
        }

        else{

            buildAssistant();

            assistantPanel.style.display =
            "block";
        }
    }
);

// =====================================
// NEXT WORKING SHIFT
// =====================================

function findNextWorkingShift(employee){

    const dates =
    Object.keys(monthRoster)
    .sort();

    let foundToday = false;

    for(const date of dates){

        if(date === currentDateKey){

            foundToday = true;
            continue;
        }

        if(!foundToday)
            continue;

        const shift =
        monthRoster[date][employee];

        if(
            shift &&
            shift !== "OFF"
        ){

            return {
                date,
                shift
            };
        }
    }

    return null;
}
// =====================================
// BUILD ASSISTANT
// =====================================

function buildAssistant(){

    assistantContent.innerHTML = "";

    employeeNames.forEach(employee => {

        const next =
        findNextWorkingShift(employee);

        const todayShift =
        todayRoster[employee] || "OFF";

        const div =
        document.createElement("div");

        div.className =
        "assistant-person";

        const photo =

        PHOTO_MAP[
            employee.toUpperCase()
        ]

        ||

        `https://ui-avatars.com/api/?background=00d9ff&color=fff&name=${encodeURIComponent(employee)}`;

        let nextShiftText =
        "No upcoming shift";

        if(next){

            const formattedDate =

            new Date(next.date)

            .toLocaleDateString(
                "en-GB",
                {
                    weekday:"short",
                    day:"2-digit",
                    month:"long"
                }
            )

            .replace(",","");

            nextShiftText =

            `Next: ${next.shift} • ${formattedDate}`;
        }

        div.innerHTML = `

            <img
                src="${photo}"
                alt="${employee}"
            >

            <div>

                <b>${employee}</b>

                <br>

                Today: ${todayShift}

                <br>

                ${nextShiftText}

            </div>

        `;

        assistantContent.appendChild(
            div
        );
    });
}
async function loadNotice(){

    try{

        const res = await fetch(
            "https://infosec-notice-api.ashrithmv.workers.dev/notice"
        );

        const notices =
        await res.json();

        const board =
        document.getElementById(
            "shiftNoticeBoard"
        );

        if(
            !Array.isArray(notices)
            ||
            notices.length === 0
        ){

            board.innerHTML = `
            <div class="notice-card">
                No active notices
            </div>
            `;

            return;
        }

        let html = "";

        notices.forEach(notice => {

            html += `

            <div class="notice-card">

                <div class="notice-sender">
                    👤 ${notice.sender}
                </div>

                <div class="notice-message">
                    ${notice.message}
                </div>

                <div class="notice-time">
                    🕒 ${new Date(
                        notice.timestamp
                    ).toLocaleString(
                        "en-IN",
                        {
                            timeZone:
                            "Asia/Kolkata"
                        }
                    )}
                </div>

            </div>
            `;
        });

        board.innerHTML = html;

    }

    catch(err){

        console.error(err);

    }
}
loadNotice();

setInterval(
    loadNotice,
    10000
);
