// ==========================================
// INFOSEC SHIFT COMMAND CENTER
// AUTO LOAD EXCEL VERSION
// CLOUDLFARE READY
// ==========================================

// ----------------------------
// DOM
// ----------------------------

const clock = document.getElementById("clock");
const todayDate = document.getElementById("todayDate");

const currentShiftEl =
document.getElementById("currentShift");

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

// ----------------------------
// PHOTOS
// ----------------------------

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

    S1:
    "07:30 AM → 03:30 PM",

    S2:
    "01:30 PM → 10:00 PM",

    S3:
    "09:30 PM → 08:00 AM"
};

// ----------------------------
// GLOBALS
// ----------------------------

let todayRoster = {};

let activeCards = [];

// ----------------------------
// CLOCK
// ----------------------------

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

// ----------------------------
// TIME HELPERS
// ----------------------------

function timeToDate(time){

    const [h,m] =
    time.split(":");

    const d =
    new Date();

    d.setHours(Number(h));
    d.setMinutes(Number(m));
    d.setSeconds(0);
    d.setMilliseconds(0);

    return d;
}

function formatRemaining(ms){

    if(ms < 0)
        ms = 0;

    const total =
    Math.floor(ms / 1000);

    const hrs =
    Math.floor(total / 3600);

    const mins =
    Math.floor(
        (total % 3600) / 60
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

// ----------------------------
// CURRENT SHIFT
// ----------------------------

function getCurrentShift(){

    const now =
    new Date();

    const mins =

        now.getHours() * 60
        +
        now.getMinutes();

    if(
        mins >= 450 &&
        mins < 930
    ){
        return "S1";
    }

    if(
        mins >= 810 &&
        mins < 1320
    ){
        return "S2";
    }

    return "S3";
}

// ----------------------------
// SHIFT END
// ----------------------------

function getShiftEnd(shift){

    const now =
    new Date();

    let end;

    if(shift === "S1"){

        end =
        timeToDate("15:30");
    }

    else if(
        shift === "S2"
    ){

        end =
        timeToDate("22:00");
    }

    else{

        end =
        timeToDate("08:00");

        if(
            now.getHours()
            >= 21
        ){

            end.setDate(
                end.getDate() + 1
            );
        }
    }

    return end;
}

// ----------------------------
// HEADER
// ----------------------------

function updateShiftHeader(){

    const shift =
    getCurrentShift();

    currentShiftEl.textContent =
    shift;

    const end =
    getShiftEnd(
        shift
    );

    shiftCountdown.textContent =
    formatRemaining(
        end - new Date()
    );
}

setInterval(
    updateShiftHeader,
    1000
);

updateShiftHeader();

// ----------------------------
// AUTO LOAD EXCEL
// ----------------------------

async function loadExcelAutomatically(){

    try{

        const response =
        await fetch(
            "shifts.xlsx"
        );

        const buffer =
        await response.arrayBuffer();

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

        processRoster(
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

// ----------------------------
// PROCESS EXCEL
// ----------------------------

function processRoster(rows){

    todayRoster = {};

    const headers =
    rows[1];

    const employees = [

        headers[1],

        headers[2],

        headers[3],

        headers[4],

        headers[5]
    ];

    const today =
    new Date();

    let todayRow =
    null;

    for(
        let i = 2;
        i < rows.length;
        i++
    ){

        const row =
        rows[i];

        if(
            !row ||
            !row[0]
        ) continue;

        const serial =
        row[0];

        if(
            typeof serial
            !== "number"
        ) continue;

        const excelDate =
        new Date(

            (
                serial
                -
                25569
            )
            *
            86400
            *
            1000
        );

        if(

            excelDate.getDate()
            ===
            today.getDate()

            &&

            excelDate.getMonth()
            ===
            today.getMonth()

        ){

            todayRow =
            row;

            break;
        }
    }

    if(!todayRow){

        console.error(
            "Today's roster not found"
        );

        return;
    }

    employees.forEach(
        (
            name,
            index
        )=>{

            todayRoster[
                name
            ] =

            String(
                todayRow[
                    index + 1
                ]
            )
            .trim()
            .toUpperCase();
        }
    );

    buildDashboard();
}

// ----------------------------
// COUNTS
// ----------------------------

function updateCounts(){

    let s1 = 0;
    let s2 = 0;
    let s3 = 0;
    let off = 0;

    Object.values(
        todayRoster
    ).forEach(
        shift=>{

            if(
                shift === "S1"
            ) s1++;

            else if(
                shift === "S2"
            ) s2++;

            else if(
                shift === "S3"
            ) s3++;

            else off++;
        }
    );

    s1Count.textContent =
    s1;

    s2Count.textContent =
    s2;

    s3Count.textContent =
    s3;

    offCount.textContent =
    off;
}

// ----------------------------
// OFF PEOPLE
// ----------------------------

function updateOffPeople(){

    const container =
    document.getElementById(
        "offPeople"
    );

    container.innerHTML = "";

    Object.entries(
        todayRoster
    ).forEach(

        ([name,shift])=>{

            if(
                shift === "OFF"
            ){

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
        }
    );
}

// ----------------------------
// BUILD DASHBOARD
// ----------------------------

function buildDashboard(){

    updateCounts();

    updateOffPeople();

    staffContainer.innerHTML =
    "";

    activeCards = [];

    const currentShift =
    getCurrentShift();

    Object.entries(
        todayRoster
    ).forEach(

        ([name,shift])=>{

            if(
                shift
                !==
                currentShift
            ) return;

            const clone =
            template.content.cloneNode(
                true
            );

            clone.querySelector(
                ".staff-name"
            ).textContent =
            name;

            clone.querySelector(
                ".staff-shift"
            ).innerHTML =

            `
            ${shift}
            <br>
            <small>
            ${SHIFT_DISPLAY[shift]}
            </small>
            `;

            const img =
            clone.querySelector(
                ".avatar"
            );

            img.src =

            PHOTO_MAP[
                String(name)
                .toUpperCase()
            ]

            ||

            `https://ui-avatars.com/api/?background=00d9ff&color=fff&name=${encodeURIComponent(name)}`;

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
// CARD UPDATE
// ----------------------------

function updateCards(){

    const currentShift =
    getCurrentShift();

    const info =
    SHIFTS[
        currentShift
    ];

    const start =
    timeToDate(
        info.start
    );

    const end =
    getShiftEnd(
        currentShift
    );

    activeCards.forEach(
        item=>{

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

            const total =
            end - start;

            const done =
            new Date() -
            start;

            let percent =
            (
                done
                /
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

            fill.style.width =
            percent + "%";

            text.textContent =
            Math.floor(
                percent
            )
            + "%";

            remain.textContent =

            formatRemaining(
                end - new Date()
            )

            +

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

function buildNextShift(){

    const current =
    getCurrentShift();

    let next;

    if(
        current === "S1"
    ){

        next = "S2";
    }

    else if(
        current === "S2"
    ){

        next = "S3";
    }

    else{

        next = "S1";
    }

    nextShiftName.textContent =
    next;

    nextShiftTime.textContent =
    SHIFT_DISPLAY[
        next
    ];

    const people =

    Object.entries(
        todayRoster
    )

    .filter(
        ([n,s])=>
        s === next
    )

    .map(
        ([n])=>n
    );

    nextShiftPeople.innerHTML =

    people.length

    ?

    people.join("<br>")

    :

    "No staff";
}

// ----------------------------
// SHIFT CHANGE AUTO REFRESH
// ----------------------------

let lastShift =
getCurrentShift();

setInterval(()=>{

    const current =
    getCurrentShift();

    if(
        current
        !==
        lastShift
    ){

        lastShift =
        current;

        buildDashboard();
    }

},30000);

// ----------------------------
// START
// ----------------------------

loadExcelAutomatically();