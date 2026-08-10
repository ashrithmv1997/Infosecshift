// ==========================================================
// INFOSEC + SOC SHIFT COMMAND CENTER
// COMPLETE script.js
// ==========================================================


// ==========================================================
// DOM ELEMENTS
// ==========================================================

const clock =
document.getElementById("clock");

const todayDate =
document.getElementById("todayDate");

const activeShiftDisplay =
document.getElementById("activeShiftDisplay");

const shiftCountdown =
document.getElementById("shiftCountdown");

const s1Count =
document.getElementById("s1Count");

const s2Count =
document.getElementById("s2Count");

const s3Count =
document.getElementById("s3Count");

const offCount =
document.getElementById("offCount");

const staffContainer =
document.getElementById("staffContainer");

const nextShiftName =
document.getElementById("nextShiftName");

const nextShiftTime =
document.getElementById("nextShiftTime");

const nextShiftPeople =
document.getElementById("nextShiftPeople");

const template =
document.getElementById("staffTemplate");


// ==========================================================
// MODALS
// ==========================================================

const calendarBtn =
document.getElementById("calendarBtn");

const offFinderBtn =
document.getElementById("offFinderBtn");

const calendarModal =
document.getElementById("calendarModal");

const offModal =
document.getElementById("offModal");

const calendarGrid =
document.getElementById("calendarGrid");

const selectedDayRoster =
document.getElementById("selectedDayRoster");

const employeeSelect =
document.getElementById("employeeSelect");

const offResults =
document.getElementById("offResults");


// ==========================================================
// SHIFT ASSISTANT
// ==========================================================

const assistantBtn =
document.getElementById("shiftAssistantBtn");

const assistantPanel =
document.getElementById("shiftAssistantPanel");

const assistantContent =
document.getElementById("assistantContent");


// ==========================================================
// TEAM SWITCHER
// ==========================================================

const dashboardTitle =
document.getElementById("dashboardTitle");

const teamSwitchButtons =
document.querySelectorAll(".team-switch-btn");


// ==========================================================
// TEAM CONFIGURATION
// ==========================================================

const TEAM_CONFIG = {

    INFOSEC: {

        label: "INFOSEC",

        file:
        "shifts.xlsx"
    },

    SOC: {

        label: "SOC",

        file:
        "SOC_Shifts.xlsx"
    }
};


// Remember last selected team

let currentTeam =

localStorage.getItem(
    "selectedShiftTeam"
)

||

"INFOSEC";


if(
    !TEAM_CONFIG[currentTeam]
){

    currentTeam =
    "INFOSEC";
}


// ==========================================================
// SHIFT DEFINITIONS
// ==========================================================

const SHIFTS = {

    S1: {

        start:
        "07:30",

        end:
        "15:30",

        label:
        "07:30 AM → 03:30 PM"
    },

    S2: {

        start:
        "13:30",

        end:
        "22:00",

        label:
        "01:30 PM → 10:00 PM"
    },

    S3: {

        start:
        "21:30",

        end:
        "08:00",

        label:
        "09:30 PM → 08:00 AM"
    }
};


// ==========================================================
// EMPLOYEE PHOTOS
// ==========================================================

const PHOTO_MAP = {

    // INFOSEC

    "ASHWIN":
    "ashwin.jpg",

    "RAHUL R G":
    "rahul.jpg",

    "HARI":
    "hari.jpg",

    "ASHRITH":
    "ashrith.jpg",

    "MANU":
    "manu.jpg",


    // SOC
    // Add their photos later if you want.
    // Until then automatic avatars are used.

    "JENSUN":
    "jensun.jpg",

    "RAHUL":
    "soc-rahul.jpg",

    "ABHIJITH":
    "abhijith.jpg",

    "GEORGE":
    "george.jpg",

    "PAVAN":
    "pavan.jpg"
};
const CONTACTS = {

    INFOSEC: {
        "ASHRITH": "918089617628",
        "RAHUL R G": "919946389225",
        "ASHWIN": "919746221620",
        "MANU": "919113843240",
        "HARI": "919847009362"
    },

    SOC: {
        "ABHIJITH": "917356497658",
        "RAHUL": "919342998819",
        "JENSUN": "919633778074",
        "GEORGE": "916282796357",
        "PAVAN": "918328103570"
    }

};

// ==========================================================
// GLOBAL DATA
// ==========================================================

let monthRoster = {};

let todayRoster = {};

let employeeNames = [];

let activeCards = [];

let currentDateKey = "";

let leaveEmployees = [];


// ==========================================================
// CLOCK
// ==========================================================

function updateClock(){

    const now =
    new Date();

    if(clock){

        clock.textContent =
        now.toLocaleTimeString(
            "en-GB"
        );
    }

    if(todayDate){

        todayDate.textContent =
        now.toLocaleDateString(

            "en-GB",

            {

                weekday:
                "long",

                day:
                "numeric",

                month:
                "long",

                year:
                "numeric"
            }
        );
    }
}


setInterval(
    updateClock,
    1000
);

updateClock();


// ==========================================================
// DATE HELPERS
// ==========================================================

function excelToDate(
    serial
){

    return new Date(

        (
            serial -
            25569
        )

        *

        86400

        *

        1000
    );
}


// ----------------------------------------------------------
// Convert whatever Excel gives us into a Date
// ----------------------------------------------------------

function rosterValueToDate(
    value
){

    if(
        value instanceof Date
    ){

        return value;
    }


    if(
        typeof value ===
        "number"
    ){

        return excelToDate(
            value
        );
    }


    const parsed =
    new Date(value);

    if(
        !isNaN(
            parsed.getTime()
        )
    ){

        return parsed;
    }


    return null;
}


// ----------------------------------------------------------
// YYYY-MM-DD
// ----------------------------------------------------------

function getDateKey(
    date
){

    const year =
    date.getFullYear();

    const month =
    String(
        date.getMonth() + 1
    )
    .padStart(
        2,
        "0"
    );

    const day =
    String(
        date.getDate()
    )
    .padStart(
        2,
        "0"
    );

    return (

        year

        +

        "-"

        +

        month

        +

        "-"

        +

        day
    );
}


// ==========================================================
// TIME HELPERS
// ==========================================================

function timeToDate(
    timeString
){

    const parts =
    timeString.split(":");

    const hour =
    Number(parts[0]);

    const minute =
    Number(parts[1]);

    const date =
    new Date();

    date.setHours(
        hour
    );

    date.setMinutes(
        minute
    );

    date.setSeconds(
        0
    );

    date.setMilliseconds(
        0
    );

    return date;
}


// ----------------------------------------------------------

function formatRemaining(
    ms
){

    if(ms < 0){

        ms = 0;
    }

    const total =
    Math.floor(
        ms / 1000
    );

    const hours =
    Math.floor(
        total / 3600
    );

    const minutes =
    Math.floor(

        (
            total %
            3600
        )

        /

        60
    );

    const seconds =
    total %
    60;

    return (

        String(hours)
        .padStart(
            2,
            "0"
        )

        +

        ":"

        +

        String(minutes)
        .padStart(
            2,
            "0"
        )

        +

        ":"

        +

        String(seconds)
        .padStart(
            2,
            "0"
        )
    );
}


// ==========================================================
// ACTIVE SHIFT ENGINE
// ==========================================================

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
    // 07:30 → 15:30

    if(

        mins >=
        450

        &&

        mins <
        930

    ){

        active.push(
            "S1"
        );
    }


    // S2
    // 13:30 → 22:00

    if(

        mins >=
        810

        &&

        mins <
        1320

    ){

        active.push(
            "S2"
        );
    }


    // S3
    // 21:30 → 08:00

    if(

        mins >=
        1290

        ||

        mins <
        480

    ){

        active.push(
            "S3"
        );
    }


    return active;
}


// ==========================================================
// NEXT SHIFT BOUNDARY
// ==========================================================

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

        time => {

            const date =
            timeToDate(
                time
            );


            if(
                date <= now
            ){

                date.setDate(

                    date.getDate()

                    +

                    1
                );
            }


            if(

                !nearest

                ||

                date <
                nearest

            ){

                nearest =
                date;
            }
        }
    );


    return nearest;
}


// ==========================================================
// HEADER
// ==========================================================

function updateHeader(){

    const shifts =
    getActiveShifts();


    if(activeShiftDisplay){

        activeShiftDisplay.textContent =

        shifts.length

        ?

        shifts.join(
            " + "
        )

        :

        "No Active Shift";
    }


    const next =
    getNextBoundary();


    if(
        shiftCountdown &&
        next
    ){

        shiftCountdown.textContent =

        formatRemaining(

            next

            -

            new Date()
        );
    }
}


setInterval(
    updateHeader,
    1000
);

updateHeader();


// ==========================================================
// TEAM SWITCH UI
// ==========================================================

function updateTeamSwitcherUI(){

    teamSwitchButtons.forEach(

        button => {

            button.classList.toggle(

                "active",

                button.dataset.team ===
                currentTeam
            );
        }
    );


    if(dashboardTitle){

        dashboardTitle.textContent =

        TEAM_CONFIG[
            currentTeam
        ].label

        +

        " SHIFT COMMAND CENTER";
    }


    document.title =

    TEAM_CONFIG[
        currentTeam
    ].label

    +

    " SHIFT COMMAND CENTER";
}


// ==========================================================
// EXCEL LOADER
// ==========================================================

async function loadExcel(){

    const fileName =

    TEAM_CONFIG[
        currentTeam
    ].file;


    try{

        console.log(

            "Loading team:",

            currentTeam,

            fileName
        );


        const response =
        await fetch(

            fileName

            +

            "?v="

            +

            Date.now(),

            {

                cache:
                "no-store"
            }
        );


        if(
            !response.ok
        ){

            throw new Error(

                "Unable to load "

                +

                fileName

                +

                " status "

                +

                response.status
            );
        }


        const buffer =
        await response.arrayBuffer();


        const workbook =
        XLSX.read(

            buffer,

            {

                type:
                "array"
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

                header:
                1
            }
        );


        processMonthRoster(
            rows
        );


        return true;
    }


    catch(error){

        console.error(

            "Roster load error:",

            error
        );


        if(staffContainer){

            staffContainer.innerHTML = `

            <div
                class="glass"
                style="
                    padding:25px;
                    border-radius:20px;
                    grid-column:1/-1;
                    text-align:center;
                "
            >

                Unable to load
                ${fileName}

            </div>
            `;
        }


        return false;
    }
}


// ==========================================================
// PROCESS ROSTER
// ==========================================================

function processMonthRoster(
    rows
){

    monthRoster = {};


    if(
        !rows ||
        rows.length < 2
    ){

        console.error(
            "Invalid roster file"
        );

        return;
    }


    // Row 2 contains employee names
    const headers =
    rows[1];


    // Automatically detect all employee columns
    // First column is Day/Date

    employeeNames = [

    String(headers[1]).trim(),
    String(headers[2]).trim(),
    String(headers[3]).trim(),
    String(headers[4]).trim(),
    String(headers[5]).trim()

];

    // Data starts from row 3

    for(

        let i = 2;

        i < rows.length;

        i++

    ){

        const row =
        rows[i];


        if(
            !row ||
            row.length === 0
        ){

            continue;
        }


        const date =
        rosterValueToDate(
            row[0]
        );


        if(!date){

            continue;
        }


        const key =
        getDateKey(
            date
        );


        monthRoster[
            key
        ] = {};


        employeeNames.forEach(

            (
                employee,
                index
            ) => {

                const rawShift =
                row[
                    index + 1
                ];


                const shift =
                String(
                    rawShift ?? "OFF"
                )

                .trim()

                .toUpperCase();


                monthRoster[
                    key
                ][
                    employee
                ] =
                shift;
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
    ]

    ||

    {};


    console.log(

        currentTeam,

        "Month Roster Loaded",

        monthRoster
    );


    console.log(

        currentTeam,

        "Today's Roster",

        todayRoster
    );


    populateEmployeeDropdown();
}


// ==========================================================
// TEAM SWITCH
// ==========================================================

async function switchTeam(
    team
){

    if(
        !TEAM_CONFIG[team]
    ){

        return;
    }


    if(
        team === currentTeam
    ){

        return;
    }


    currentTeam =
    team;


    localStorage.setItem(

        "selectedShiftTeam",

        currentTeam
    );


    updateTeamSwitcherUI();


    document.body.classList.add(
        "team-switching"
    );


    // Close assistant during change

    if(assistantPanel){

        assistantPanel.style.display =
        "none";
    }


    // Close Calendar

    if(calendarModal){

        calendarModal.classList.remove(
            "show"
        );
    }


    // Close OFF finder

    if(offModal){

        offModal.classList.remove(
            "show"
        );
    }


    // Clear old data

    monthRoster = {};

    todayRoster = {};

    employeeNames = [];

    activeCards = [];


    if(staffContainer){

        staffContainer.innerHTML = `

        <div
            class="glass"
            style="
                padding:30px;
                border-radius:20px;
                text-align:center;
                grid-column:1/-1;
            "
        >

            Loading
            ${currentTeam}
            roster...

        </div>
        `;
    }


    if(nextShiftName){

        nextShiftName.textContent =
        "--";
    }


    if(nextShiftTime){

        nextShiftTime.textContent =
        "Loading...";
    }


    if(nextShiftPeople){

        nextShiftPeople.textContent =
        "Loading...";
    }


    const offPeople =
    document.getElementById(
        "offPeople"
    );


    if(offPeople){

        offPeople.textContent =
        "Loading...";
    }


    const loaded =
    await loadExcel();


    if(loaded){

        initializeDashboard();
    }


    document.body.classList.remove(
        "team-switching"
    );


    console.log(

        "Team switched to",

        currentTeam
    );
}


// ----------------------------------------------------------

teamSwitchButtons.forEach(

    button => {

        button.addEventListener(

            "click",

            () => {

                switchTeam(
                    button.dataset.team
                );
            }
        );
    }
);


// ==========================================================
// EMPLOYEE DROPDOWN
// ==========================================================

function populateEmployeeDropdown(){

    if(!employeeSelect){

        return;
    }


    employeeSelect.innerHTML = `

        <option value="">
            Select Employee
        </option>
    `;


    employeeNames.forEach(

        employee => {

            const option =
            document.createElement(
                "option"
            );


            option.value =
            employee;


            option.textContent =
            employee;


            employeeSelect.appendChild(
                option
            );
        }
    );
}


// ==========================================================
// SHIFT START / END
// ==========================================================

function getShiftStart(
    shift
){

    const now =
    new Date();


    let start;


    if(
        shift === "S1"
    ){

        start =
        timeToDate(
            "07:30"
        );
    }


    else if(
        shift === "S2"
    ){

        start =
        timeToDate(
            "13:30"
        );
    }


    else{

        start =
        timeToDate(
            "21:30"
        );


        // After midnight S3 started yesterday

        if(
            now.getHours() < 8
        ){

            start.setDate(

                start.getDate()

                -

                1
            );
        }
    }


    return start;
}


// ----------------------------------------------------------

function getShiftEnd(
    shift
){

    const now =
    new Date();


    let end;


    if(
        shift === "S1"
    ){

        end =
        timeToDate(
            "15:30"
        );
    }


    else if(
        shift === "S2"
    ){

        end =
        timeToDate(
            "22:00"
        );
    }


    else{

        end =
        timeToDate(
            "08:00"
        );


        // S3 starting tonight ends tomorrow

        if(
            now.getHours() >= 21
        ){

            end.setDate(

                end.getDate()

                +

                1
            );
        }
    }


    return end;
}


// ==========================================================
// COUNTERS
// ==========================================================

function updateCounts(){

    let s1 = 0;

    let s2 = 0;

    let s3 = 0;

    let off = 0;


    Object.values(
        todayRoster
    )

    .forEach(

        shift => {

            if(
                shift === "S1"
            ){

                s1++;
            }

            else if(
                shift === "S2"
            ){

                s2++;
            }

            else if(
                shift === "S3"
            ){

                s3++;
            }

            else{

                off++;
            }
        }
    );


    if(s1Count){

        s1Count.textContent =
        s1;
    }


    if(s2Count){

        s2Count.textContent =
        s2;
    }


    if(s3Count){

        s3Count.textContent =
        s3;
    }


    if(offCount){

        offCount.textContent =
        off;
    }
}


// ==========================================================
// OFF TODAY
// ==========================================================

function updateOffPeople(){

    const container =
    document.getElementById(
        "offPeople"
    );


    if(!container){

        return;
    }


    container.innerHTML =
    "";


    const offStaff =

    Object.entries(
        todayRoster
    )

    .filter(

        (
            [
                name,
                shift
            ]
        ) =>

        shift ===
        "OFF"
    );


    if(
        offStaff.length === 0
    ){

        container.textContent =
        "No one OFF today";

        return;
    }


    offStaff.forEach(

        (
            [
                name
            ]
        ) => {

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
    );
}


// ==========================================================
// NEXT OFF
// ==========================================================

function findNextOff(
    employee
){

    const dates =

    Object.keys(
        monthRoster
    )

    .sort();


    for(
        const date of dates
    ){

        if(
            date <
            currentDateKey
        ){

            continue;
        }


        const roster =
        monthRoster[
            date
        ];


        if(
            !roster
        ){

            continue;
        }


        if(
            roster[
                employee
            ] === "OFF"
        ){

            const d =
            new Date(
                date +
                "T12:00:00"
            );


            return d.toLocaleDateString(

                "en-GB",

                {

                    weekday:
                    "short",

                    day:
                    "numeric",

                    month:
                    "short"
                }
            );
        }
    }


    return "No OFF";
}


// ==========================================================
// CURRENTLY WORKING
// IMPORTANT: S3 AFTER MIDNIGHT USES YESTERDAY
// ==========================================================

function getEmployeesCurrentlyWorking(){

    const now =
    new Date();


    const mins =

    now.getHours()

    *

    60

    +

    now.getMinutes();


    const activeShifts =
    getActiveShifts();


    const result = [];


    // ------------------------------------------------------
    // 00:00 → 08:00
    // S3 belongs to yesterday
    // ------------------------------------------------------

    if(
        mins <
        480
    ){

        const yesterday =
        new Date();


        yesterday.setDate(

            yesterday.getDate()

            -

            1
        );


        const yesterdayKey =
        getDateKey(
            yesterday
        );


        const yesterdayRoster =

        monthRoster[
            yesterdayKey
        ]

        ||

        {};


        Object.entries(
            yesterdayRoster
        )

        .forEach(

            (
                [
                    name,
                    shift
                ]
            ) => {

                if(
                    shift ===
                    "S3"
                ){

                    result.push({

                        name,

                        shift
                    });
                }
            }
        );
    }


    // ------------------------------------------------------
    // Today's S1/S2/S3
    // ------------------------------------------------------

    Object.entries(
        todayRoster
    )

    .forEach(

        (
            [
                name,
                shift
            ]
        ) => {

            if(
                !activeShifts.includes(
                    shift
                )
            ){

                return;
            }


            // During 00:00 → 08:00,
            // don't show today's S3.
            // Yesterday's S3 is still working.

            if(

                mins <
                480

                &&

                shift ===
                "S3"

            ){

                return;
            }


            result.push({

                name,

                shift
            });
        }
    );


    return result;
}


// ==========================================================
// DASHBOARD BUILD
// ==========================================================

function buildDashboard(){

    updateCounts();

    updateOffPeople();


    if(staffContainer){

        staffContainer.innerHTML =
        "";
    }


    activeCards = [];


    const staff =
    getEmployeesCurrentlyWorking();


    if(
        staff.length === 0 &&
        staffContainer
    ){

        staffContainer.innerHTML = `

        <div
            class="glass"
            style="
                padding:25px;
                border-radius:20px;
                text-align:center;
                grid-column:1/-1;
            "
        >

            No staff currently working

        </div>
        `;
    }


    staff.forEach(

        employee => {

            createStaffCard(

                employee.name,

                employee.shift
            );
        }
    );


    buildUpcomingShift();

    updateCards();
}


// ==========================================================
// STAFF CARD
// ==========================================================

function createStaffCard(
    name,
    shift
){

    if(
        !template ||
        !staffContainer
    ){
        return;
    }


    // ==========================================
    // CLONE STAFF CARD
    // ==========================================

    const clone =
    template.content.cloneNode(
        true
    );


    const avatar =
    clone.querySelector(
        ".avatar"
    );


    const nameElement =
    clone.querySelector(
        ".staff-name"
    );


    const shiftElement =
    clone.querySelector(
        ".staff-shift"
    );


    const nextOffElement =
    clone.querySelector(
        ".next-off"
    );


    const card =
    clone.querySelector(
        ".staff-card"
    );


    // ==========================================
    // NAME
    // ==========================================

    if(nameElement){

        nameElement.textContent =
        name;
    }


    // ==========================================
    // SHIFT
    // ==========================================

    if(shiftElement){

        shiftElement.innerHTML = `

            ${shift}

            <br>

            <small>
                ${SHIFTS[shift].label}
            </small>

        `;
    }


    const key =
    name
    .toUpperCase()
    .trim();


    // ==========================================
    // PHOTO
    // ==========================================

    if(avatar){

        avatar.src =

        PHOTO_MAP[key]

        ||

        `https://ui-avatars.com/api/?background=00d9ff&color=fff&name=${encodeURIComponent(name)}`;


        avatar.onerror = () => {

            avatar.onerror =
            null;

            avatar.src =

            `https://ui-avatars.com/api/?background=00d9ff&color=fff&name=${encodeURIComponent(name)}`;

        };
    }


    // ==========================================
    // NEXT OFF
    // ==========================================

    if(nextOffElement){

        nextOffElement.innerHTML = `

            Next OFF:

            <br>

            ${findNextOff(name)}

        `;
    }


    // ==========================================
    // CONTACT NUMBERS
    // ==========================================

    const CONTACTS = {

        INFOSEC: {

            "ASHRITH":
            "918089617628",

            "RAHUL R G":
            "919946389225",

            "ASHWIN":
            "919746221620",

            "MANU":
            "919113843240",

            "HARI":
            "919847009362"

        },


        SOC: {

            "ABHIJITH":
            "917356497658",

            "RAHUL":
            "919342998819",

            "JENSUN":
            "919633778074",

            "GEORGE":
            "916282796357",

            "PAVAN":
            "918328103570"

        }

    };


    const phone =

    CONTACTS[
        currentTeam
    ]?.[
        key
    ];


    // ==========================================
    // CALL + WHATSAPP BUTTONS
    // ==========================================

    if(
        phone &&
        card
    ){

        const contactWrap =
        document.createElement(
            "div"
        );


        contactWrap.className =
        "staff-contact-actions";


        contactWrap.innerHTML = `

            <a
                class="staff-contact-btn call-btn"
                href="tel:+${phone}"
                title="Call ${name}"
                aria-label="Call ${name}"
            >

                <span class="call-icon">
                    ☎
                </span>

                <span>
                    Call
                </span>

            </a>


            <a
                class="staff-contact-btn whatsapp-btn"
                href="https://wa.me/${phone}"
                target="_blank"
                rel="noopener noreferrer"
                title="WhatsApp ${name}"
                aria-label="WhatsApp ${name}"
            >

                <svg
                    class="whatsapp-svg"
                    viewBox="0 0 32 32"
                    aria-hidden="true"
                >

                    <path
                        fill="currentColor"
                        d="
                        M16.04 3
                        C8.86 3 3.02 8.82 3.02 15.98
                        c0 2.29.6 4.52 1.73 6.49
                        L3 29
                        l6.7-1.75
                        a13 13 0 0 0 6.33 1.61
                        h.01
                        c7.17 0 13.01-5.82 13.01-12.98
                        C29.05 8.82 23.21 3 16.04 3
                        Z

                        m0 23.67
                        h-.01
                        a10.78 10.78 0 0 1-5.5-1.5
                        l-.39-.23
                        l-3.97 1.04
                        l1.06-3.86
                        l-.25-.4
                        a10.7 10.7 0 0 1-1.65-5.74
                        c0-5.95 4.85-10.79 10.82-10.79
                        2.89 0 5.6 1.12 7.64 3.16
                        a10.7 10.7 0 0 1 3.17 7.63
                        c-.01 5.95-4.86 10.79-10.82 10.79
                        Z

                        m5.93-8.08
                        c-.32-.16-1.92-.95-2.22-1.06
                        -.3-.11-.51-.16-.73.16
                        -.22.33-.84 1.06-1.03 1.28
                        -.19.22-.38.24-.7.08
                        -.33-.16-1.37-.5-2.61-1.61
                        a9.78 9.78 0 0 1-1.81-2.25
                        c-.19-.33-.02-.5.14-.66
                        .15-.15.33-.38.49-.57
                        .16-.19.22-.33.33-.54
                        .11-.22.05-.41-.03-.57
                        -.08-.16-.73-1.76-1-2.41
                        -.26-.63-.53-.55-.73-.56
                        h-.62
                        c-.22 0-.57.08-.87.41
                        -.3.33-1.14 1.12-1.14 2.72
                        0 1.6 1.17 3.15 1.33 3.37
                        .16.22 2.3 3.51 5.57 4.92
                        .78.34 1.39.54 1.86.69
                        .78.25 1.49.21 2.05.13
                        .63-.09 1.92-.79 2.19-1.55
                        .27-.76.27-1.41.19-1.55
                        -.08-.14-.3-.22-.62-.38
                        Z
                        "
                    />

                </svg>

                <span>
                    WhatsApp
                </span>

            </a>

        `;


        // Put buttons BEFORE progress bar

        const progressWrap =
        card.querySelector(
            ".progress-wrap"
        );


        if(progressWrap){

            progressWrap.before(
                contactWrap
            );

        }else if(nextOffElement){

            nextOffElement.before(
                contactWrap
            );

        }else{

            card.appendChild(
                contactWrap
            );
        }
    }


    // ==========================================
    // ON LEAVE
    // ==========================================

    if(

        card

        &&

        leaveEmployees.includes(
            key
        )

    ){

        card.classList.add(
            "on-leave"
        );


        const badge =
        document.createElement(
            "div"
        );


        badge.className =
        "leave-badge";


        badge.textContent =
        "🔴 ON LEAVE";


        card.appendChild(
            badge
        );
    }


    // ==========================================
    // ADD CARD
    // ==========================================

    staffContainer.appendChild(
        clone
    );


    activeCards.push({

        card,

        shift

    });

}

// ==========================================================
// PROGRESS BARS
// ==========================================================

function updateCards(){

    const now =
    new Date();


    activeCards.forEach(

        item => {

            const start =
            getShiftStart(
                item.shift
            );


            const end =
            getShiftEnd(
                item.shift
            );


            const total =
            end - start;


            const elapsed =
            now - start;


            let percent =

            (
                elapsed /
                total
            )

            *

            100;


            percent =
            Math.max(

                0,

                Math.min(
                    100,
                    percent
                )
            );


            const fill =
            item.card
            ?.querySelector(
                ".progress-fill"
            );


            const text =
            item.card
            ?.querySelector(
                ".progress-text"
            );


            const remaining =
            item.card
            ?.querySelector(
                ".remaining-time"
            );


            if(fill){

                fill.style.width =
                percent + "%";
            }


            if(text){

                text.textContent =

                Math.floor(
                    percent
                )

                +

                "% Complete";
            }


            if(remaining){

                remaining.textContent =

                formatRemaining(

                    end

                    -

                    now
                )

                +

                " Remaining";
            }
        }
    );
}


setInterval(
    updateCards,
    1000
);


// ==========================================================
// UPCOMING SHIFT
// ==========================================================

function buildUpcomingShift(){

    if(
        !nextShiftName ||
        !nextShiftTime ||
        !nextShiftPeople
    ){

        return;
    }


    const now =
    new Date();


    const mins =

    now.getHours()

    *

    60

    +

    now.getMinutes();


    let nextShift;

    let rosterDate;


    // ------------------------------------------------------
    // Before S1 starts
    // ------------------------------------------------------

    if(
        mins <
        450
    ){

        nextShift =
        "S1";


        rosterDate =
        new Date();
    }


    // ------------------------------------------------------
    // S1 started, next S2
    // ------------------------------------------------------

    else if(
        mins <
        810
    ){

        nextShift =
        "S2";


        rosterDate =
        new Date();
    }


    // ------------------------------------------------------
    // S2 started, next S3
    // ------------------------------------------------------

    else if(
        mins <
        1290
    ){

        nextShift =
        "S3";


        rosterDate =
        new Date();
    }


    // ------------------------------------------------------
    // S3 started - next S1 tomorrow
    // ------------------------------------------------------

    else{

        nextShift =
        "S1";


        rosterDate =
        new Date();


        rosterDate.setDate(

            rosterDate.getDate()

            +

            1
        );
    }


    const rosterDateKey =
    getDateKey(
        rosterDate
    );


    const roster =

    monthRoster[
        rosterDateKey
    ]

    ||

    {};


    const people =

    Object.entries(
        roster
    )

    .filter(

        (
            [
                name,
                shift
            ]
        ) =>

        shift ===
        nextShift
    )

    .map(

        (
            [
                name
            ]
        ) =>

        name
    );


    nextShiftName.textContent =
    nextShift;


    nextShiftTime.textContent =
    SHIFTS[
        nextShift
    ].label;


    nextShiftPeople.innerHTML =

    people.length

    ?

    people.join(
        "<br>"
    )

    :

    "No Staff";
}


// ==========================================================
// CALENDAR
// ==========================================================

function buildCalendar(){

    if(!calendarGrid){

        return;
    }


    calendarGrid.innerHTML =
    "";


    const dates =

    Object.keys(
        monthRoster
    )

    .sort();


    dates.forEach(

        dateKey => {

            const date =
            new Date(
                dateKey +
                "T12:00:00"
            );


            const div =
            document.createElement(
                "div"
            );


            div.className =
            "calendar-day";


            div.innerHTML = `

                <strong>
                    ${date.getDate()}
                </strong>

                <br>

                <small>

                    ${date.toLocaleDateString(
                        "en-GB",
                        {
                            weekday:
                            "short"
                        }
                    )}

                </small>
            `;


            div.addEventListener(

                "click",

                () => {

                    showRosterForDate(
                        dateKey
                    );
                }
            );


            calendarGrid.appendChild(
                div
            );
        }
    );
}


// ==========================================================
// SELECTED DATE ROSTER
// ==========================================================

function showRosterForDate(
    dateKey
){

    if(!selectedDayRoster){

        return;
    }


    const roster =
    monthRoster[
        dateKey
    ];


    if(!roster){

        selectedDayRoster.innerHTML =
        "No roster found";

        return;
    }


    const s1 = [];

    const s2 = [];

    const s3 = [];

    const off = [];


    Object.entries(
        roster
    )

    .forEach(

        (
            [
                name,
                shift
            ]
        ) => {

            if(
                shift === "S1"
            ){

                s1.push(
                    name
                );
            }


            else if(
                shift === "S2"
            ){

                s2.push(
                    name
                );
            }


            else if(
                shift === "S3"
            ){

                s3.push(
                    name
                );
            }


            else{

                off.push(
                    name
                );
            }
        }
    );


    const date =
    new Date(
        dateKey +
        "T12:00:00"
    );


    selectedDayRoster.innerHTML = `

        <h3>

            ${date.toLocaleDateString(

                "en-GB",

                {

                    weekday:
                    "long",

                    day:
                    "numeric",

                    month:
                    "long"
                }
            )}

        </h3>

        <br>

        <b>S1</b>

        <br>

        ${s1.join(", ") || "-"}

        <br><br>

        <b>S2</b>

        <br>

        ${s2.join(", ") || "-"}

        <br><br>

        <b>S3</b>

        <br>

        ${s3.join(", ") || "-"}

        <br><br>

        <b>OFF</b>

        <br>

        ${off.join(", ") || "-"}
    `;
}


// ==========================================================
// CALENDAR OPEN
// ==========================================================

if(calendarBtn){

    calendarBtn.addEventListener(

        "click",

        () => {

            buildCalendar();


            if(selectedDayRoster){

                selectedDayRoster.innerHTML = `

                    <h3>
                        📅 Select a Date
                    </h3>

                    <p>
                        Click any date below to view roster
                    </p>
                `;
            }


            calendarModal
            ?.classList
            .add(
                "show"
            );
        }
    );
}


// ==========================================================
// OFF FINDER OPEN
// ==========================================================

if(offFinderBtn){

    offFinderBtn.addEventListener(

        "click",

        () => {

            populateEmployeeDropdown();


            offModal
            ?.classList
            .add(
                "show"
            );
        }
    );
}


// ==========================================================
// MODAL CLOSE BUTTONS
// ==========================================================

document

.querySelectorAll(
    ".close-btn"
)

.forEach(

    button => {

        button.addEventListener(

            "click",

            () => {

                calendarModal
                ?.classList
                .remove(
                    "show"
                );


                offModal
                ?.classList
                .remove(
                    "show"
                );
            }
        );
    }
);


// ==========================================================
// CLICK OUTSIDE MODAL
// ==========================================================

window.addEventListener(

    "click",

    event => {

        if(
            event.target ===
            calendarModal
        ){

            calendarModal.classList.remove(
                "show"
            );
        }


        if(
            event.target ===
            offModal
        ){

            offModal.classList.remove(
                "show"
            );
        }
    }
);


// ==========================================================
// OFF FINDER CHANGE
// ==========================================================

if(employeeSelect){

    employeeSelect.addEventListener(

        "change",

        showEmployeeOffs
    );
}


// ==========================================================
// SHOW EMPLOYEE OFF DAYS
// ==========================================================

function showEmployeeOffs(){

    if(
        !employeeSelect ||
        !offResults
    ){

        return;
    }


    const employee =
    employeeSelect.value;


    if(!employee){

        offResults.innerHTML =
        "Select an employee";

        return;
    }


    const offDays = [];


    Object.keys(
        monthRoster
    )

    .sort()

    .forEach(

        dateKey => {

            const roster =
            monthRoster[
                dateKey
            ];


            if(
                roster?.[
                    employee
                ] !== "OFF"
            ){

                return;
            }


            const date =
            new Date(
                dateKey +
                "T12:00:00"
            );


            offDays.push(`

                <div class="off-person">

                    ${date.toLocaleDateString(

                        "en-GB",

                        {

                            weekday:
                            "short",

                            day:
                            "numeric",

                            month:
                            "short"
                        }
                    )}

                </div>
            `);
        }
    );


    offResults.innerHTML = `

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


// ==========================================================
// SHIFT ASSISTANT
// ==========================================================

if(
    assistantBtn &&
    assistantPanel &&
    assistantContent
){

    assistantBtn.addEventListener(

        "click",

        () => {

            if(
                assistantPanel.style.display ===
                "block"
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
}


// ==========================================================
// FIND NEXT WORKING SHIFT
// ==========================================================

function findNextWorkingShift(
    employee
){

    const dates =

    Object.keys(
        monthRoster
    )

    .sort();


    let foundToday =
    false;


    for(
        const date of dates
    ){

        if(
            date ===
            currentDateKey
        ){

            foundToday =
            true;

            continue;
        }


        if(
            !foundToday
        ){

            continue;
        }


        const shift =

        monthRoster[
            date
        ]?.[
            employee
        ];


        if(

            shift

            &&

            shift !==
            "OFF"

        ){

            return {

                date,

                shift
            };
        }
    }


    return null;
}


// ==========================================================
// BUILD SHIFT ASSISTANT
// ==========================================================

function buildAssistant(){

    if(!assistantContent){

        return;
    }


    assistantContent.innerHTML =
    "";


    employeeNames.forEach(

        employee => {

            const next =
            findNextWorkingShift(
                employee
            );


            const todayShift =

            todayRoster[
                employee
            ]

            ||

            "OFF";


            const div =
            document.createElement(
                "div"
            );


            div.className =
            "assistant-person";


            const key =
            employee
            .toUpperCase()
            .trim();


            const photo =

            PHOTO_MAP[
                key
            ]

            ||

            `https://ui-avatars.com/api/?background=00d9ff&color=fff&name=${encodeURIComponent(employee)}`;


            let nextShiftText =
            "No upcoming shift";


            if(next){

                const formattedDate =

                new Date(
                    next.date +
                    "T12:00:00"
                )

                .toLocaleDateString(

                    "en-GB",

                    {

                        weekday:
                        "short",

                        day:
                        "2-digit",

                        month:
                        "long"
                    }
                )

                .replace(
                    ",",
                    ""
                );


                nextShiftText =

                `Next: ${next.shift} • ${formattedDate}`;
            }


            div.innerHTML = `

                <img
                    src="${photo}"
                    alt="${employee}"
                    onerror="
                        this.onerror=null;
                        this.src='https://ui-avatars.com/api/?background=00d9ff&color=fff&name=${encodeURIComponent(employee)}'
                    "
                >

                <div>

                    <b>
                        ${employee}
                    </b>

                    <br>

                    Today:
                    ${todayShift}

                    <br>

                    ${nextShiftText}

                </div>
            `;


            assistantContent.appendChild(
                div
            );
        }
    );
}


// ==========================================================
// LEAVE STATUS
// ==========================================================

async function loadLeaveStatus(){

    try{

        const response =
        await fetch(

            "https://infosec-notice-api.ashrithmv.workers.dev/leave"

            +

            "?v="

            +

            Date.now(),

            {

                cache:
                "no-store"
            }
        );


        if(
            !response.ok
        ){

            return;
        }


        const data =
        await response.json();


        if(
            !Array.isArray(data)
        ){

            return;
        }


        const previous =
        JSON.stringify(
            leaveEmployees
        );


        leaveEmployees =

        data

        .map(

            item =>

            String(
                item.employee ||
                ""
            )

            .trim()

            .toUpperCase()
        )

        .filter(
            Boolean
        );


        // Only rebuild when leave status changed

        if(

            JSON.stringify(
                leaveEmployees
            )

            !==

            previous

        ){

            buildDashboard();
        }
    }


    catch(error){

        console.error(

            "Leave Load Error",

            error
        );
    }
}


// ==========================================================
// SHIFT NOTICES
// ==========================================================

async function loadNotice(){

    const board =
    document.getElementById(
        "shiftNoticeBoard"
    );


    if(!board){

        return;
    }


    try{

        const response =
        await fetch(

            "https://infosec-notice-api.ashrithmv.workers.dev/notice"

            +

            "?v="

            +

            Date.now(),

            {

                cache:
                "no-store"
            }
        );


        if(
            !response.ok
        ){

            return;
        }


        const notices =
        await response.json();


        if(

            !Array.isArray(
                notices
            )

            ||

            notices.length ===
            0

        ){

            board.innerHTML = `

                <div class="notice-card">

                    No active notices

                </div>
            `;


            return;
        }


        let html =
        "";


        notices.forEach(

            notice => {

                const sender =
                notice.sender ||
                "Unknown";


                const message =
                notice.message ||
                "";


                let formattedTime =
                "";


                if(
                    notice.timestamp
                ){

                    const date =
                    new Date(
                        notice.timestamp
                    );


                    if(
                        !isNaN(
                            date.getTime()
                        )
                    ){

                        formattedTime =
                        date.toLocaleString(

                            "en-IN",

                            {

                                timeZone:
                                "Asia/Kolkata",

                                weekday:
                                "short",

                                day:
                                "2-digit",

                                month:
                                "short",

                                hour:
                                "2-digit",

                                minute:
                                "2-digit"
                            }
                        );
                    }
                }


                html += `

                    <div class="notice-card">

                        <div class="notice-sender">

                            👤 ${sender}

                        </div>

                        <div class="notice-message">

                            ${message}

                        </div>

                        <div class="notice-time">

                            ${
                                formattedTime

                                ?

                                "🕒 " +
                                formattedTime

                                :

                                ""
                            }

                        </div>

                    </div>
                `;
            }
        );


        board.innerHTML =
        html;
    }


    catch(error){

        console.error(

            "Notice Load Error",

            error
        );
    }
}


// ==========================================================
// DATE CHANGE
// ==========================================================

function checkDateChange(){

    const newKey =
    getDateKey(
        new Date()
    );


    if(
        newKey ===
        currentDateKey
    ){

        return;
    }


    currentDateKey =
    newKey;


    todayRoster =

    monthRoster[
        newKey
    ]

    ||

    {};


    buildDashboard();


    if(
        assistantPanel &&
        assistantPanel.style.display ===
        "block"
    ){

        buildAssistant();
    }
}


setInterval(
    checkDateChange,
    30000
);


// ==========================================================
// ACTIVE SHIFT CHANGE WATCHER
// ==========================================================

let lastShiftState =

JSON.stringify(
    getActiveShifts()
);


setInterval(

    () => {

        const currentState =

        JSON.stringify(
            getActiveShifts()
        );


        if(
            currentState !==
            lastShiftState
        ){

            lastShiftState =
            currentState;


            buildDashboard();


            if(
                assistantPanel &&
                assistantPanel.style.display ===
                "block"
            ){

                buildAssistant();
            }
        }
    },

    15000
);


// ==========================================================
// INITIALIZE DASHBOARD
// ==========================================================

function initializeDashboard(){

    updateHeader();

    updateCounts();

    updateOffPeople();

    buildDashboard();

    buildUpcomingShift();

    populateEmployeeDropdown();

    updateCards();


    if(
        assistantPanel &&
        assistantPanel.style.display ===
        "block"
    ){

        buildAssistant();
    }
}


// ==========================================================
// START SYSTEM
// ==========================================================

async function startSystem(){

    updateTeamSwitcherUI();


    const loaded =
    await loadExcel();


    if(loaded){

        initializeDashboard();
    }


    loadNotice();

    loadLeaveStatus();


    console.log(

        "SHIFT COMMAND CENTER READY",

        currentTeam
    );
}


startSystem();


// ==========================================================
// ROSTER AUTO REFRESH EVERY 5 MINUTES
// ==========================================================

setInterval(

    async () => {

        const loaded =
        await loadExcel();


        if(loaded){

            initializeDashboard();
        }


        console.log(

            currentTeam,

            "Roster refreshed"
        );
    },

    300000
);


// ==========================================================
// NOTICE AUTO REFRESH
// ==========================================================

setInterval(
    loadNotice,
    10000
);


// ==========================================================
// LEAVE AUTO REFRESH
// ==========================================================

setInterval(
    loadLeaveStatus,
    5000
);


// ==========================================================
// WHEN APP / TAB BECOMES ACTIVE
// ==========================================================

document.addEventListener(

    "visibilitychange",

    async () => {

        if(
            document.hidden
        ){

            return;
        }


        console.log(
            "App resumed"
        );


        const loaded =
        await loadExcel();


        if(loaded){

            initializeDashboard();
        }


        loadNotice();

        loadLeaveStatus();
    }
);


// ==========================================================
// END
// ==========================================================
