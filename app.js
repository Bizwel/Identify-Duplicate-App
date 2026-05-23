document.getElementById("processBtn")
  .addEventListener("click", processFile);

function processFile() {

  const fileInput =
    document.getElementById("fileInput");

  const status =
    document.getElementById("status");

  if (!fileInput.files.length) {

    status.innerHTML =
      "Please upload a file.";

    return;
  }

  const file = fileInput.files[0];

  const reader = new FileReader();

  reader.onload = function(e) {

    try {

      const data =
        new Uint8Array(e.target.result);

      // =========================
      // READ WORKBOOK
      // =========================

      const workbook = XLSX.read(data, {
        type: "array"
      });

      const firstSheetName =
        workbook.SheetNames[0];

      const worksheet =
        workbook.Sheets[firstSheetName];

      const jsonData =
        XLSX.utils.sheet_to_json(worksheet);

      if (!jsonData.length) {

        status.innerHTML =
          "No data found.";

        return;
      }

      // =========================
      // NORMALIZE STATUS
      // =========================

      function normalizeStatus(status) {

        status = (status || "")
          .toString()
          .trim()
          .toLowerCase();

        // Invoice variations
        if (
          status.includes("invoice")
        ) {
          return "invoice";
        }

        // Partial Refund variations
        if (
          status.includes("partial refund")
        ) {
          return "partial refund";
        }

        // Saved
        if (
          status.includes("saved")
        ) {
          return "saved";
        }

        return status;
      }

      // =========================
      // GROUP BY TICKET NO
      // =========================

      const groupedTickets = {};

      jsonData.forEach(row => {

        const ticketNo =
          row["Ticket No"];

        if (!ticketNo) return;

        if (!groupedTickets[ticketNo]) {

          groupedTickets[ticketNo] = [];
        }

        groupedTickets[ticketNo].push(row);

      });

      // =========================
      // PROCESS RECORDS
      // =========================

      const duplicates = [];
      const excludedRecords = [];

      Object.keys(groupedTickets)
        .forEach(ticketNo => {

        const records =
          groupedTickets[ticketNo];

        // Only process duplicates
        if (records.length > 1) {

          // IMPORTANT:
          // Support BOTH Status and Folder Status

          const statuses = records.map(r => {

            const rawStatus =
              r["Status"] ||
              r["Folder Status"] ||
              "";

            return normalizeStatus(rawStatus);

          });

          const hasPartialRefund =
            statuses.includes(
              "partial refund"
            );

          const hasInvoice =
            statuses.includes(
              "invoice"
            );

          // =========================
          // EXCLUDE VALID CASES
          // =========================

          if (
            hasPartialRefund &&
            hasInvoice
          ) {

            excludedRecords.push(
              ...records
            );

            return;
          }

          // =========================
          // TRUE DUPLICATES
          // =========================

          duplicates.push(
            ...records
          );

        }

      });

      // =========================
      // CREATE OUTPUT FILE
      // =========================

      const outputWorkbook =
        XLSX.utils.book_new();

      // Original Data
      const originalSheet =
        XLSX.utils.json_to_sheet(
          jsonData
        );

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        originalSheet,
        "Original Data"
      );

      // Duplicate Tickets
      const duplicateSheet =
        XLSX.utils.json_to_sheet(
          duplicates
        );

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        duplicateSheet,
        "Duplicate Tickets"
      );

      // Excluded Records
      const excludedSheet =
        XLSX.utils.json_to_sheet(
          excludedRecords
        );

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        excludedSheet,
        "Excluded Records"
      );

      // =========================
      // EXPORT
      // =========================

      XLSX.writeFile(
        outputWorkbook,
        "Duplicate_Ticket_Report.xlsx"
      );

      // =========================
      // UI STATUS
      // =========================

      status.innerHTML = `
        <strong>Processing Complete.</strong>
        <br><br>

        Total Records:
        ${jsonData.length}

        <br>

        Duplicate Records:
        ${duplicates.length}

        <br>

        Excluded Valid Refund Cases:
        ${excludedRecords.length}
      `;

    } catch (error) {

      console.error(error);

      status.innerHTML =
        "Error processing file.";

    }

  };

  reader.readAsArrayBuffer(file);

}
