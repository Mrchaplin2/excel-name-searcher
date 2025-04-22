function searchNames() {
  Excel.run(async (context) => {
    const nameInput = document.getElementById("nameInput").value.trim().toLowerCase();
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange().load("values");

    await context.sync();

    const rows = usedRange.values;
    const matches = [];

    for (let row of rows) {
      for (let cell of row) {
        if (cell.toString().toLowerCase().includes(nameInput)) {
          matches.push(row);
          break;
        }
      }
    }

    const resultsDiv = document.getElementById("results");
    if (matches.length === 0) {
      resultsDiv.innerHTML = "<p>No match found.</p>";
    } else {
      const table = matches.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join('')}</tr>`).join('');
      resultsDiv.innerHTML = `<table>${table}</table>`;
    }
  }).catch(console.error);
}
