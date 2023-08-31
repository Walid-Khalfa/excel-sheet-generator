document.addEventListener("DOMContentLoaded", function () {
  const generateButton = document.querySelector("button[data-action='generate']");
  const exportButton = document.querySelector("button[data-action='export']");

  generateButton.addEventListener("click", function () {
    const rowsInput = document.querySelector(".rows");
    const columnsInput = document.querySelector(".columns");
    const rows = parseInt(rowsInput.value);
    const columns = parseInt(columnsInput.value);

    if (isNaN(rows) || isNaN(columns) || rows <= 0 || columns <= 0) {
      Swal.fire("Error", "Please enter valid numbers for rows and columns.", "error");
    } else {
      generateTable(rows, columns);
    }
  });

  exportButton.addEventListener("click", function () {
    const table = document.querySelector(".table table");

    if (!table || table.rows.length <= 0) {
      Swal.fire("Error", "No table to export.", "error");
    } else {
      ExportToExcel("xlsx");
    }
  });
});

function generateTable(rows, columns) {
  const sheetBody = document.querySelector(".sheet-body");
  sheetBody.innerHTML = "";

  for (let i = 0; i < rows; i++) {
    const row = document.createElement("tr");

    for (let j = 0; j < columns; j++) {
      const cell = document.createElement("td");
      cell.textContent = `Row ${i + 1}, Col ${j + 1}`;
      row.appendChild(cell);
    }

    sheetBody.appendChild(row);
  }
}

const ExportToExcel = (fileType) => {
  const table = document.querySelector(".table table");

  if (!table || table.rows.length <= 0 || table.rows[0].cells.length <= 0) {
    Swal.fire("Error", "No table to export.", "error");
    return;
  }

  const wb = XLSX.utils.table_to_book(table, { sheet: "SheetJS" });
  const wbout = XLSX.write(wb, { bookType: fileType, bookSST: true, type: "array" });

  saveAs(new Blob([wbout], { type: "application/octet-stream" }), "sheet." + fileType);
};
