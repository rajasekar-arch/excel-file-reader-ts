
import { getExcelRowCount } from "excel-file-reader-ts"; // Do npm install excel-file-reader
import * as path from "path"; // Do npm install path

async function runTest() {
  const filePath = path.join(__dirname, "sample.xlsx"); // Assumes sample.xlsx is in the same directory

  try {
    // Test with default (first) sheet
    const rowCountDefault = await getExcelRowCount(filePath);
    console.log(`Number of rows in default sheet: ${rowCountDefault}`); // Expected: the count of excel sheet

    // Test with a specific sheet name (if your Excel has multiple sheets)
    // const rowCountSheet2 = await getExcelRowCount(filePath, 'Sheet2');
    // console.log(`Number of rows in 'Sheet2': ${rowCountSheet2}`);

    // Test non-existent file (should throw error)
    try {
      await getExcelRowCount("non-existent-file.xlsx");
    } catch (error: any) {
      console.error(`Error for non-existent file (expected): ${error.message}`);
    }

    // Test non-existent sheet (should throw error if sheetName is provided)
    try {
      await getExcelRowCount(filePath, "NonExistentSheet");
    } catch (error: any) {
      console.error(
        `Error for non-existent sheet (expected): ${error.message}`
      );
    }
  } catch (error: any) {
    console.error(`An unexpected error occurred: ${error.message}`);
  }
}

runTest();
