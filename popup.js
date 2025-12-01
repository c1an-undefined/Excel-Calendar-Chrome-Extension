let allEvents = [];
let filteredEvents = [];

document.getElementById("parse-btn").addEventListener("click", async () => {
    setParsedStatus("Parsing…");
    setButtonState(false, null); // Disable everything while parsing

    const fileInput = document.getElementById("file")
    const file = fileInput.files[0]
    if (!file) {
        console.log("No file selected");
        return;
    }

    const data = await file.arrayBuffer();

    const workbook = XLSX.read(data)

    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]

    const rows = XLSX.utils.sheet_to_json(sheet)

    const lecturerFilter = document.getElementById("username").value.trim();

    rows.forEach(row => {
        allEvents.push(parseEvent(row))
    });

    filteredEvents = allEvents.filter(event =>
        event.lecturer.toLowerCase().includes(lecturerFilter.toLowerCase())
    );

    if (filteredEvents.length === 0) {
        setParsedStatus("File parsed — no matching events.");
        setButtonState(true, null);
    } else {
        setParsedStatus(`File parsed — ${filteredEvents.length} events found.`);
        setButtonState(true, null);
    }

    populateDropdown(filteredEvents);

    document.getElementById("event-dropdown").addEventListener("change", () => {
        const val = document.getElementById("event-dropdown").value;
        setButtonState(true, val);
    });


    document.getElementById("add-one").addEventListener("click", async () => {
        const index = document.getElementById("event-dropdown").value;
        if (!index) return;

        const event = filteredEvents[index];

        console.log("Adding event:", event);

        const result = await addEventToCalendar(event);

        console.log("Calendar result:", result);
        alert("Added to Calendar!");
    });

    document.getElementById("add-all").addEventListener("click", async () => {
        const status = document.getElementById("status");
        status.innerHTML = "Adding events...";

        const added = [];
        const skipped = [];

        for (let i = 0; i < filteredEvents.length; i++) {
            const event = filteredEvents[i];

            const duplicate = await checkDuplicate(event);
            if (duplicate) {
                skipped.push(event.title);
                continue;
            }

            await addEventToCalendar(event);
            added.push(event.title);
        }

        status.innerHTML = `
            <strong>Done!</strong><br>
            Added: ${added.length} <br>
            Skipped duplicates: ${skipped.length}
        `;
    });
})

async function checkDuplicate(event) {
    return new Promise((resolve) => {
        chrome.runtime.sendMessage({ type: "GET_TOKEN" }, async (response) => {
            const token = response.token;

            const start = event.start.toISOString();
            const end = event.end.toISOString();

            const url =
                `https://www.googleapis.com/calendar/v3/calendars/primary/events?` +
                `timeMin=${start}&timeMax=${end}&q=${encodeURIComponent(event.title)}`;

            const res = await fetch(url, {
                headers: { Authorization: "Bearer " + token }
            });

            const data = await res.json();

            resolve(data.items && data.items.length > 0);
        });
    });
}


function populateDropdown(events) {
    const dropdown = document.getElementById("event-dropdown");
    dropdown.innerHTML = '<option value="">-- Select an event --</option>';

    const seen = new Set();

    events.forEach((event, index) => {
        const id = `${event.title} - ${event.start.toISOString()} - ${event.end.toISOString()}`;

        if (seen.has(id)) return;  // skip duplicates
        seen.add(id);

        const option = document.createElement("option");
        option.value = index;  // store index in array
        option.textContent = formatOption(event);
        dropdown.appendChild(option);
    });
}

function formatOption(event) {
  const start = new Date(event.start);
  const end = new Date(event.end);

  const day = String(start.getDate()).padStart(2, '0');
  const month = String(start.getMonth() + 1).padStart(2, '0');

  const startTime = start.toLocaleTimeString('en-IE', {
    hour: '2-digit',
    minute: '2-digit'
  });

  const endTime = end.toLocaleTimeString('en-IE', {
    hour: '2-digit',
    minute: '2-digit'
  });

  return `${event.title} — ${day}/${month} · ${startTime}–${endTime}`;
}

async function addEventToCalendar(event) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage({ type: "GET_TOKEN" }, async (response) => {
      if (response.error) {
        reject(response.error);
        return;
      }

      const token = response.token;

      const googleEvent = {
        summary: event.title,
        description: `${event.description}`,
        location: `${event.location}`,
        start: {
            dateTime: event.start.toISOString(),
            timeZone: "Europe/Dublin"
        },
        end: {
            dateTime: event.end.toISOString(),
            timeZone: "Europe/Dublin"
        }
      };


      try {
        const res = await fetch(
          "https://www.googleapis.com/calendar/v3/calendars/primary/events",
          {
            method: "POST",
            headers: {
              Authorization: "Bearer " + token,
              "Content-Type": "application/json"
            },
            body: JSON.stringify(googleEvent)
          }
        );

        const data = await res.json();
        resolve(data);
      } catch (err) {
        reject(err);
      }
    });
  });
}

function parseEvent(row) {
    const dateStr = getColumn(row, "date");
    const timeStr = getColumn(row, "time");
    const title = getColumn(row, "module");
    const lecturer = getColumn(row, "lecturer");
    const description = getColumn(row, "content");
    const location = getColumn(row, "room");

    const baseDate = parseDate(dateStr);
    const { start, end } = parseTimeRange(timeStr, baseDate);

    return {
        title,
        lecturer,
        description,
        location,
        start,
        end
    };
}


function getColumn(row, target) {
    const targetLower = target.toLowerCase();
    for (const key of Object.keys(row)) {
        if (key.toLowerCase().includes(targetLower)) {
            return row[key];
        }
    }
    return null;
}

function parseTimeRange(timeString, baseDate) {
    if (!timeString) {
        console.log("Missing time value, defaulting to 1 hour at 09:00:", timeString);
        timeString = "09:00"; 
    }

    timeString = String(timeString).trim();  // <-- FIX (force string so .trim never fails)

    const parts = timeString.split("-");

    if (parts.length === 2) {
        const [sh, sm] = parts[0].split(":").map(Number);
        const [eh, em] = parts[1].split(":").map(Number);

        const start = new Date(baseDate);
        const end = new Date(baseDate);

        start.setHours(sh, sm || 0, 0);
        end.setHours(eh, em || 0, 0);

        return { start, end };
    }

    // SINGLE TIME CASE — "09:00 Induction"
    const timeMatch = timeString.match(/(\d{1,2}):(\d{2})/);
    if (!timeMatch) {
        console.log("Unrecognized time, defaulting to 09:00:", timeString);
        const start = new Date(baseDate);
        const end = new Date(baseDate);

        start.setHours(9, 0, 0);
        end.setHours(10, 0, 0);

        return { start, end };
    }

    const [_, h, m] = timeMatch.map(Number);

    const start = new Date(baseDate);
    const end = new Date(baseDate);

    start.setHours(h, m, 0);
    end.setHours(h + 1, m, 0);

    return { start, end };
}


function parseDate(input) {
    // Case 1: Excel serial number (e.g., 45973)
    if (typeof input === "number") {
        return new Date((input - 25569) * 86400 * 1000);
    }

    // Case 2: dd/mm/yyyy string
    if (typeof input === "string") {
        const parts = input.trim().split("/");
        if (parts.length === 3) {
            const [day, month, year] = parts.map(Number);
            return new Date(year, month - 1, day);
        }
    }

    console.log("Unrecognized date format:", input);
    return new Date(); // fallback to avoid crashes
}

function setParsedStatus(text) {
    document.getElementById("parse-status").textContent = text;
}

function setButtonState(parsed, dropdownValue) {
    const addOne = document.getElementById("add-one");
    const addAll = document.getElementById("add-all");

    // Disable both if file not parsed
    if (!parsed) {
        addOne.disabled = true;
        addAll.disabled = true;
        return;
    }

    // Add All enabled only if we have events
    addAll.disabled = filteredEvents.length === 0;

    // Add One enabled only when event selected
    addOne.disabled = !dropdownValue;
}

window.addEventListener("DOMContentLoaded", () => {
    setButtonState(false, null);
});

