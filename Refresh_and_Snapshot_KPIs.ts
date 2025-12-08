/**
 * Script Name: Refresh_and_Snapshot_KPIs
 * Description: Triggers a Power Query refresh, waits for completion, 
 * and snapshots the latest data from Column C into a new historical Column E.
 */https://github.com/Tan-Zhi-Yuan/power-automate-library/tree/main

async function main(workbook: ExcelScript.Workbook) {
  // 1. Refresh All Data Connections (Power Query)
  // We use 'await' to ensure the refresh triggers before we try to copy data.
  console.log("Starting Data Refresh...");
  await workbook.refreshAllDataConnections();
  console.log("Data Refresh Triggered.");

  // 2. Define Sheet and Ranges
  const sheetName = "Summary";
  const summarySheet = workbook.getWorksheet(sheetName);

  // Safety Check: Ensure sheet exists
  if (!summarySheet) {
    console.log(`Error: Worksheet '${sheetName}' not found.`);
    return;
  }

  // 3. Insert New Column at E (Shifts existing history to the right)
  // This preserves historical data in F, G, H, etc.
  summarySheet.getRange("E:E").insert(ExcelScript.InsertShiftDirection.right);

  // 4. Snapshot Data (Copy Values Only)
  // Copying from the "Live" column (C) to the "Snapshot" column (E)
  let sourceRange = summarySheet.getRange("C6:C69");
  let destinationRange = summarySheet.getRange("E6:E69");

  destinationRange.copyFrom(
    sourceRange, 
    ExcelScript.RangeCopyType.values, // Values only (break formulas)
    false, // Skip Blanks
    false  // Transpose
  );

  // 5. Formatting
  summarySheet.getRange("E:E").getFormat().setColumnWidth(85.5);
  
  console.log("Snapshot Complete.");
}
