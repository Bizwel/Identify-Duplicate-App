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

    // Group rows by Ticket No
    const groupedTickets = {};

    jsonData.forEach(row => {

      const ticketNo = row["Ticket No"];

      if (!ticketNo) return;

      if (!groupedTickets[ticketNo]) {
        groupedTickets[ticketNo] = [];
      }

      groupedTickets[ticketNo].push(row);

    });

    const duplicates = [];

    // Process each Ticket No
    Object.keys(groupedTickets).forEach(ticketNo => {

      const records = groupedTickets[ticketNo];

      // Only interested in duplicates
      if (records.length > 1) {

        // Get all Folder Status values
        const statuses = records.map(r =>
          (r["Folder Status"] || "").toString().trim().toLowerCase()
        );

        const hasPartialRefund = statuses.includes("partial refund");
        const hasInvoice = statuses.includes("invoice");

        // EXCLUDE if Partial Refund + Invoice combination exists
        if (hasPartialRefund && hasInvoice) {
          return;
        }

        // Otherwise include all duplicate rows
        duplicates.push(...records);
      }

    });

    // Create output workbook
    const newWorkbook = XLSX.utils.book_new();

    // Original Data Sheet
    const originalSheet = XLSX.utils.json_to_sheet(jsonData);

    XLSX.utils.book_append_sheet(
      newWorkbook,
      originalSheet,
      "Original Data"
    );

    // Duplicate Sheet
    const duplicateSheet = XLSX.utils.json_to_sheet(duplicates);

    XLSX.utils.book_append_sheet(
      newWorkbook,
      duplicateSheet,
      "Duplicate Tickets"
    );

    // Export
    XLSX.writeFile(
      newWorkbook,
      "Filtered_Duplicate_Report.xlsx"
    );

    status.innerHTML =
      `Processing complete. ${duplicates.length} duplicate rows found after exclusions.`;

  };

  reader.readAsArrayBuffer(file);
}
