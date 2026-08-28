import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = fileURLToPath(new URL("../tests/fixtures/", import.meta.url));
await fs.mkdir(outputDir, { recursive: true });

const headers = [
  "Type",
  "Item",
  "Quantity",
  "Description",
  "Desired Profit",
  "Net Cost",
  "Unit Profit Amount",
  "Unit Price",
  "Extended Cost",
  "Extended Profit",
  "Net Price",
  "Category",
];

async function createExport(fileName, title, rows) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Export");
  sheet.getRange("A1:L1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1:L1").format = {
    fill: "#17324D",
    font: { bold: true, color: "#FFFFFF", size: 14 },
    rowHeight: 24,
  };
  sheet.getRange("A3:L3").values = [headers];
  sheet.getRange("A3:L3").format = {
    fill: "#D88A3D",
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
  };
  const lastRow = 3 + rows.length;
  sheet.getRange(`A4:L${lastRow}`).values = rows;
  sheet.getRange(`C4:C${lastRow}`).format.numberFormat = "0.00";
  sheet.getRange(`E4:K${lastRow}`).format.numberFormat = "$#,##0.00";
  sheet.getRange(`A3:L${lastRow}`).format.borders = {
    top: { style: "thin", color: "#D8DEE5" },
    bottom: { style: "thin", color: "#D8DEE5" },
    left: { style: "thin", color: "#D8DEE5" },
    right: { style: "thin", color: "#D8DEE5" },
  };
  sheet.getRange("A:A").format.columnWidth = 13;
  sheet.getRange("B:B").format.columnWidth = 15;
  sheet.getRange("C:C").format.columnWidth = 10;
  sheet.getRange("D:D").format.columnWidth = 38;
  sheet.getRange("E:K").format.columnWidth = 14;
  sheet.getRange("L:L").format.columnWidth = 18;
  sheet.freezePanes.freezeRows(3);

  const inspection = await workbook.inspect({
    kind: "table",
    range: `Export!A1:L${lastRow}`,
    include: "values,formulas",
    tableMaxRows: 20,
    tableMaxCols: 12,
  });
  if (!inspection.ndjson.includes("Net Price")) {
    throw new Error(`Fixture inspection failed for ${fileName}`);
  }
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(path.join(outputDir, fileName));
}

await createExport("floor-1-accessories.xlsx", "Synthetic Floor 1 Accessories", [
  ["Item", "AC-100", 2, "Grab Bar", 0.3, 100, 50, 150, 200, 100, 300, "Accessories"],
  ["Freight", "FREIGHT-IN", 1, "Inbound freight", 0, 40, 0, 40, 40, 0, 40, "Internal"],
]);

await createExport("floor-2-accessories.xlsx", "Synthetic Floor 2 Accessories", [
  ["Item", "AC-200", 1, "Paper Towel Dispenser", 0.25, 200, 80, 280, 200, 80, 280, "Accessories"],
  ["Freight", "FREIGHT-IN", 1, "KOHLER 14380-CP Soap Dispenser", 0.2, 150, 50, 200, 150, 50, 200, "Accessories"],
  ["Labor", "INSTALL", 1, "Installation allowance", 0, 120, 0, 120, 120, 0, 120, "Internal"],
]);

await createExport("floor-2-partitions.xlsx", "Synthetic Floor 2 Partitions", [
  ["Item", "PART-PANEL", 4, "Partition panels", 0.25, 500, 200, 700, 2000, 800, 2800, "Partitions"],
  ["Labor", "INSTALL", 1, "Installation allowance", 0, 700, 0, 700, 700, 0, 700, "Internal"],
]);
