const EXCEL_EXPORT_FIELDS = [
  { label: "ID Aset", key: "idAset" },
  { label: "Nama Alat", key: "namaAlat" },
  { label: "Merek", key: "merek" },
  { label: "Kategori", key: "kategori" },
  { label: "Jumlah", key: "jumlah" },
  { label: "Spesifikasi", key: "spesifikasi" },
  { label: "Kondisi", key: "kondisi" }
];

function escapeHtml(text) {
  return text
    ?.toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;") ?? "";
}

function buildExcelHtml(rows) {
  const rowHtml = rows
    .map(
      row => `<tr>${row.map(cell => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`
    )
    .join("");

  return `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8">
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
   <x:ExcelWorksheets>
     <x:ExcelWorksheet>
       <x:Name>Inventaris</x:Name>
       <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
     </x:ExcelWorksheet>
   </x:ExcelWorksheets>
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<body>
<table border="1" cellpadding="4" cellspacing="0">
  ${rowHtml}
</table>
</body>
</html>`;
}

function createExcelDataTable(data) {
  const headerRow = EXCEL_EXPORT_FIELDS.map(field => field.label);
  const dataRows = data.map(item =>
    EXCEL_EXPORT_FIELDS.map(field => {
      const value = item[field.key];
      if (field.key === 'jumlah') return Number(value) || 0;
      return value || "";
    })
  );
  return [headerRow, ...dataRows];
}

function downloadFile(content, fileName, contentType) {
  const blob = new Blob([content], { type: contentType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

function exportDataToExcel(data) {
  if (!Array.isArray(data) || data.length === 0) {
    alert("Tidak ada data yang dapat diekspor.");
    return;
  }

  const tableData = createExcelDataTable(data);
  const html = buildExcelHtml(tableData);
  downloadFile(html, "inventaris-studio-data.xls", "application/vnd.ms-excel;charset=utf-8;");
}
