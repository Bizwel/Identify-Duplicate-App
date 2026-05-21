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

    // Detect duplicates
    const ticketMap = {};
    const duplicates = [];

    jsonData.forEach(row => {

      const ticketNo = row["Ticket No"];

      if (!ticketNo) return;

      if (ticketMap[ticketNo]) {

        // Add first occurrence if not already added
        if (!ticketMap[ticketNo].added) {
          duplicates.push(ticketMap[ticketNo].row);
          ticketMap[ticketNo].added = true;
        }

        duplicates.push(row);

      } else {

        ticketMap[ticketNo] = {
          row: row,
          added: false
        };
      }

    });

    // Create new workbook
    const newWorkbook = XLSX.utils.book_new();

    // Original Sheet
    const originalSheet = XLSX.utils.json_to_sheet(jsonData);
    XLSX.utils.book_append_sheet(newWorkbook, originalSheet, "Original Data");

    // Duplicate Sheet
    const duplicateSheet = XLSX.utils.json_to_sheet(duplicates);
    XLSX.utils.book_append_sheet(newWorkbook, duplicateSheet, "Duplicate Tickets");

    // Download file
    XLSX.writeFile(newWorkbook, "Duplicate_Ticket_Report.xlsx");

    status.innerHTML =
      `Processing complete. ${duplicates.length} duplicate records found.`;

  };

  reader.readAsArrayBuffer(file);
}
