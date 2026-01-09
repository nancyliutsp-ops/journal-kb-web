let table;

document.getElementById("excelFile").addEventListener("change", handleFile);

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    renderTable(json);
  };
  reader.readAsArrayBuffer(e.target.files[0]);
}

function renderTable(data) {
  if (table) table.destroy();

  const columns = Object.keys(data[0]).map(k => ({
    title: k,
    data: k
  }));

  table = $('#journalTable').DataTable({
    data,
    columns,
    pageLength: 25
  });
}
