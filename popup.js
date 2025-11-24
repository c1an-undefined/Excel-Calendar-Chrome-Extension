document.getElementById("parse-btn").addEventListener("click", async () => {
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
    let formattedRows = []

    const lecturerFilter = document.getElementById("username").value.trim();

    rows.forEach(row => {
        formattedRows.push(parseEvent(row))
    });

    const filtered = formattedRows.filter(event =>
        event.lecturer.toLowerCase().includes(lecturerFilter.toLowerCase())
    );

    populateDropdown(filtered);

    document.getElementById("add-one").addEventListener("click", () => {
        const dropdown = document.getElementById("event-dropdown");
        const index = dropdown.value;

        if (index === "") {
            console.log("No event selected");
            return;
        }

        const event = filtered[index];
        console.log("Adding ONE event:", event);

        // next step will be calling background.js or calendar API
    });
    document.getElementById("add-all").addEventListener("click", () => {
        console.log("Adding ALL events:", filtered);

        // next: send to background.js (we can do that together)
    });

})

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
        console.log(option)
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



function excelDateToJSDate(serial) {
    return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

function parseEvent(row) {
    // Convert Excel serial to JS date (midnight)
    const day = excelDateToJSDate(row.Date);

    // Parse time range
    const [startStr, endStr] = row.Time.split("-");
    
    const start = new Date(day);
    const end = new Date(day);

    const [sh, sm] = startStr.split(":").map(Number);
    const [eh, em] = endStr.split(":").map(Number);

    start.setHours(sh, sm, 0);
    end.setHours(eh, em, 0);

    return {
        title: row.Lecture,
        lecturer: row.Lecturer,
        start,
        end
    };
}
