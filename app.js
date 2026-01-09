let table;

// 页面加载完成后自动执行
window.onload = loadExcelFromRepo;

async function loadExcelFromRepo() {
  const response = await fetch("data/journals.xlsx");
  const arrayBuffer = await response.arrayBuffer();

  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet);

  renderTable(json);
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
