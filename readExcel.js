const XLSX = require("xlsx");

module.exports = function readExcel(filePath) {
  const wb = XLSX.readFile(filePath);
  const result = {};

  wb.SheetNames.forEach((sheetName) => {
    const ws = wb.Sheets[sheetName];

    const raw = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: "",
    });

    if (raw.length < 2) return;

    // ===== 1. หา header row ที่เหมาะสม =====
    let headerRowIndex = -1;

    for (let i = 0; i < raw.length; i++) {
      const row = raw[i];
      const validCells = row.filter(
        (c) => typeof c === "string" && c.trim() !== ""
      );

      if (validCells.length >= 3) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) return;

    const headerRow = raw[headerRowIndex];

    // ⭐ จำนวน column จริง = ช่องที่ header ไม่ว่าง
    const columnCount = headerRow.filter(
      (h) => String(h).trim() !== ""
    ).length;

    // ===== 2. สร้าง headers เฉพาะที่ไม่เกิน =====
    const headers = headerRow
      .slice(0, columnCount)
      .map((h) => String(h).trim());

    // ===== 3. เอา data เฉพาะ column ที่ไม่เกิน =====
    const rows = raw.slice(headerRowIndex + 1).map((rowArr) => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = rowArr[i] ?? "";
      });
      return obj;
    });

    result[sheetName] = {
      title: sheetName,
      headers,
      rows,
      excelRowOffset: headerRowIndex + 2,
    };
  });

  return result;
};
