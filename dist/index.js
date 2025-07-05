"use strict";
// src/index.ts
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.getExcelRowCount = getExcelRowCount;
exports.getExcelHeaders = getExcelHeaders;
exports.isExcelColumnPopulated = isExcelColumnPopulated;
const XLSX = __importStar(require("xlsx"));
const fs = __importStar(require("fs")); // Node.js file system module
/**
 * Reads an Excel file and returns the number of rows in a specified sheet.
 *
 * @param filePath The path to the Excel file (.xlsx, .xls).
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with the number of rows, or rejects with an error.
 */
async function getExcelRowCount(filePath, sheetName) {
    // Check if the file exists
    if (!fs.existsSync(filePath)) {
        throw new Error(`File not found at: ${filePath}`);
    }
    try {
        // Read the workbook
        const workbook = XLSX.readFile(filePath);
        let targetSheetName;
        // Determine the sheet to use
        if (sheetName) {
            targetSheetName = sheetName;
            if (!workbook.SheetNames.includes(targetSheetName)) {
                throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
            }
        }
        else {
            // Use the first sheet if no sheetName is provided
            if (workbook.SheetNames.length === 0) {
                throw new Error("No sheets found in the Excel file.");
            }
            targetSheetName = workbook.SheetNames[0];
        }
        const worksheet = workbook.Sheets[targetSheetName];
        // Determine the range of the worksheet to get the number of rows
        // XLSX.utils.decode_range converts a range string (e.g., "A1:C5") into an object {s: {c, r}, e: {c, r}}
        // where 's' is start and 'e' is end, and 'r' is row index.
        if (!worksheet || !worksheet['!ref']) {
            // If there's no data or no defined range, consider it 0 rows
            return 0;
        }
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        // The number of rows is (end_row_index - start_row_index + 1)
        // +1 is needed because row indices are 0-based.
        // If the sheet is empty but has a ref like "A1", range.e.r will be 0 and range.s.r will be 0.
        // So (0 - 0 + 1) = 1 row. If we want to count actual data rows, we might adjust this.
        // For simplicity, we'll count based on the highest row index in the range.
        // If you want to exclude headers, you'd subtract 1 from the total.
        const rowCount = range.e.r + 1;
        return rowCount;
    }
    catch (error) {
        // Re-throw with a more descriptive message
        throw new Error(`Failed to read Excel file or get row count: ${error}`);
    }
}
/**
 * Reads an Excel file and returns the headers (first row) of a specified sheet.
 *
 * @param filePath The path to the Excel file (.xlsx, .xls).
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with an array of header strings, or rejects with an error.
 */
async function getExcelHeaders(filePath, sheetName) {
    if (!fs.existsSync(filePath)) {
        throw new Error(`File not found at: ${filePath}`);
    }
    try {
        const workbook = XLSX.readFile(filePath);
        const targetSheetName = sheetName || workbook.SheetNames[0];
        if (!workbook.SheetNames.includes(targetSheetName)) {
            throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
        }
        const worksheet = workbook.Sheets[targetSheetName];
        if (!worksheet || !worksheet['!ref']) {
            return []; // Empty sheet or no defined range, thus no headers
        }
        // Convert the first row to an array of arrays, then take the first array as headers
        const headers = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 0, raw: false });
        if (headers.length > 0) {
            // Filter out any null/undefined/empty string headers that might result from empty cells in the first row
            return headers[0].filter(header => typeof header === 'string' && header.trim() !== '');
        }
        else {
            return []; // No headers found
        }
    }
    catch (error) {
        throw new Error(`Failed to get headers from Excel file: ${error.message}`);
    }
}
/**
 * Reads an Excel file and checks if a specific column (identified by its header name)
 * contains any non-empty values in its data rows.
 *
 * @param filePath The path to the Excel file (.xlsx, .xls).
 * @param headerName The exact name of the header column to check.
 * @param sheetName (Optional) The name of the sheet to read. If not provided, the first sheet will be used.
 * @returns A Promise that resolves with `true` if the column has at least one non-empty value, `false` otherwise.
 * Rejects with an error if the file or sheet is not found, or if the header name does not exist.
 */
async function isExcelColumnPopulated(filePath, headerName, sheetName) {
    if (!fs.existsSync(filePath)) {
        throw new Error(`File not found at: ${filePath}`);
    }
    try {
        const workbook = XLSX.readFile(filePath);
        const targetSheetName = sheetName || workbook.SheetNames[0];
        if (!workbook.SheetNames.includes(targetSheetName)) {
            throw new Error(`Sheet '${targetSheetName}' not found in the Excel file.`);
        }
        const worksheet = workbook.Sheets[targetSheetName];
        if (!worksheet || !worksheet['!ref']) {
            return false; // Empty sheet, no data, so column is not populated
        }
        // Convert sheet to an array of objects, where keys are headers
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });
        if (jsonData.length === 0) {
            return false; // No data rows, so column is not populated
        }
        // Check if the header exists in the first data row (which represents the header row in this context)
        const headers = Object.keys(jsonData[0]);
        if (!headers.includes(headerName)) {
            throw new Error(`Header '${headerName}' not found in the Excel sheet.`);
        }
        // Iterate through data rows (skipping the header row, as sheet_to_json already handles it)
        for (const row of jsonData) {
            const value = row[headerName];
            // Check if the value is not null, undefined, and not an empty string after trimming
            if (value !== null && value !== undefined && String(value).trim() !== '') {
                return true; // Found at least one populated cell in the column
            }
        }
        return false; // No populated cells found in the column
    }
    catch (error) {
        throw new Error(`Failed to check column population in Excel file: ${error}`);
    }
}
