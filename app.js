document.getElementById("processBtn").addEventListener("click", processFile);

function processFile() {

  const fileInput = document.getElementById("fileInput");
  const status = document.getElementById("status");

  if (!fileInput.files.length) {
    status.innerHTML = "Please upload a file.";
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {

    const data = new Uint8Array(e.target.result);

    const workbook = XLSX.read(data, { type: "array" });

    const firstSheetName = workbook.SheetNames[0];

    const worksheet = workbook.Sheets[firstSheetName];

    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    if (!jsonData.length) {
      status.innerHTML = "No data found.";
      return;
    }

    // =========================
    // STEP 1 — COUNT TICKETS
    // =========================

    const ticketCount = {};

    jsonData.forEach(row => {

      const ticketNo = row["Ticket No"];

      if (!ticketNo) return;

      ticketCount[ticketNo] =
        (ticketCount[ticketNo] || 0) + 1;
    });

    // =========================
    // STEP 2 — FILTER DUPLICATES
    // =========================

    const duplicates = [];

    jsonData.forEach(row => {

      const ticketNo = row["Ticket No"];
      const folderStatus =
        String(row["Folder Status"] || "").trim().toLowerCase();

      const invoice =
        String(row["Invoice"] || "").trim();

      // Skip empty Ticket No
      if (!ticketNo) return;

      // Check duplicate
      const isDuplicate = ticketCount[ticketNo] > 1;

      // Exclusion Rule
      const excludeRecord =
        folderStatus === "partial refund" &&
        invoice !== "";

      // Add only valid duplicates
      if (isDuplicate && !excludeRecord) {
        duplicates.push(row);
      }

    });

    // =========================
    // STEP 3 — CREATE OUTPUT
    // =========================

    const newWorkbook = XLSX.utils.book_new();

    // Original Data Sheet
    const originalSheet =
      XLSX.utils.json_to_sheet(jsonData);

    XLSX.utils.book_append_sheet(
      newWorkbook,
      originalSheet,
      "Original Data"
    );

    // Duplicate Sheet
    const duplicateSheet =
      XLSX.utils.json_to_sheet(duplicates);

    XLSX.utils.book_append_sheet(
      newWorkbook,
      duplicateSheet,
      "Duplicate Tickets"
    );

    // =========================
    // STEP 4 — DOWNLOAD
    // =========================

    XLSX.writeFile(
      newWorkbook,
      "Filtered_Duplicate_Tickets.xlsx"
    );

    status.innerHTML =
      `Completed. ${duplicates.length} duplicate records exported.`;

  };

  reader.readAsArrayBuffer(file);
}
