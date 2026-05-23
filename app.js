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

    try {

      const data = new Uint8Array(e.target.result);

      // Read workbook
      const workbook = XLSX.read(data, {
        type: "array"
      });

      // Get first sheet
      const firstSheetName = workbook.SheetNames[0];

      const worksheet = workbook.Sheets[firstSheetName];

      // Convert to JSON
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      if (!jsonData.length) {
        status.innerHTML = "No data found in file.";
        return;
      }

      // ============================
      // NORMALIZE STATUS FUNCTION
      // ============================

      function normalizeStatus(status) {

        status = (status || "")
          .toString()
          .trim()
          .toLowerCase();

        // Normalize all invoice variations
        if (
          status.includes("invoice") ||
          status.includes("invoiced")
        ) {
          return "invoice";
        }

        // Normalize all partial refund variations
        if (status.includes("partial refund")) {
          return "partial refund";
        }

        // Normalize saved
        if (status.includes("saved")) {
          return "saved";
        }

        // Return original normalized text
        return status;
      }

      // ============================
      // GROUP RECORDS BY TICKET NO
      // ============================

      const groupedTickets = {};

      jsonData.forEach(row => {

        const ticketNo = row["Ticket No"];

        if (!ticketNo) return;

        if (!groupedTickets[ticketNo]) {
          groupedTickets[ticketNo] = [];
        }

        groupedTickets[ticketNo].push(row);

      });

      // ============================
      // PROCESS DUPLICATES
      // ============================

      const duplicates = [];
      const excludedRecords = [];

      Object.keys(groupedTickets).forEach(ticketNo => {

        const records = groupedTickets[ticketNo];

        // Only process duplicate ticket numbers
        if (records.length > 1) {

          // Normalize statuses
          const statuses = records.map(r =>
            normalizeStatus(r["Folder Status"])
          );

          // Check conditions
          const hasPartialRefund =
            statuses.includes("partial refund");

          const hasInvoice =
            statuses.includes("invoice");

          // =========================================
          // EXCLUDE VALID PARTIAL REFUND + INVOICE
          // =========================================

          if (hasPartialRefund && hasInvoice) {

            excludedRecords.push(...records);

            return;
          }

          // =========================================
          // INCLUDE REAL DUPLICATES
          // =========================================

          duplicates.push(...records);

        }

      });

      // ============================
      // CREATE OUTPUT WORKBOOK
      // ============================

      const outputWorkbook = XLSX.utils.book_new();

      // Original Data Sheet
      const originalSheet =
        XLSX.utils.json_to_sheet(jsonData);

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        originalSheet,
        "Original Data"
      );

      // Duplicate Sheet
      const duplicateSheet =
        XLSX.utils.json_to_sheet(duplicates);

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        duplicateSheet,
        "Duplicate Tickets"
      );

      // Excluded Sheet
      const excludedSheet =
        XLSX.utils.json_to_sheet(excludedRecords);

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        excludedSheet,
        "Excluded Records"
      );

      // ============================
      // EXPORT FILE
      // ============================

      XLSX.writeFile(
        outputWorkbook,
        "Duplicate_Ticket_Report.xlsx"
      );

      // ============================
      // STATUS MESSAGE
      // ============================

      status.innerHTML = `
        Processing Complete.<br><br>
        Total Records: ${jsonData.length}<br>
        Duplicate Records: ${duplicates.length}<br>
        Excluded Valid Refund Cases: ${excludedRecords.length}
      `;

    } catch (error) {

      console.error(error);

      status.innerHTML =
        "An error occurred while processing the file.";

    }

  };

  reader.readAsArrayBuffer(file);
}
