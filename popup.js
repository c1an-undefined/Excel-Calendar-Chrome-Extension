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

    console.log(rows)
})