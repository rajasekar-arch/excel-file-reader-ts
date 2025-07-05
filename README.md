# excel-file-reader-ts
A lightweight and efficient TypeScript utility for reading Excel files (.xlsx, .xls) in Node.js environments. This package provides functions to easily get the total number of rows, extract sheet headers, and check if specific columns contain data.

# Table of Contents
    Description
    Features
    Installation
    Usage
        Get Total Row Count
        Get Headers
        Check if Column is Populated
    API Reference
    Error Handling
    Contributing
    License

# Description
excel-file-reader-ts simplifies common Excel file reading tasks for Node.js applications. Built with TypeScript, it offers type safety and clear interfaces for interacting with .xlsx and .xls files. Whether you need to quickly count records, validate header structures, or ensure data completeness in specific columns, this package provides the necessary tools.

It leverages the powerful xlsx library under the hood, abstracting away its complexities for these specific use cases.

# Features
    Get Total Row Count: Quickly determine the total number of rows in an Excel sheet, including headers.

    Extract Headers: Retrieve the names of all columns from the first row of a specified sheet.

    Check Column Population: Verify if a column (identified by its header name) contains any non-empty values in its data rows.

    TypeScript Support: Fully typed for a better developer experience.

    Promise-based API: All functions return Promises for easy asynchronous handling.

    Error Handling: Robust error messages for common issues like file not found or sheet not found.

# Installation
To install excel-file-reader-ts in your project, use npm or yarn:

npm install excel-file-reader-ts

# or
yarn add excel-file-reader-ts

#Usage
Here's how to use the functions provided by excel-file-reader-ts.

First, import the necessary functions:

// For TypeScript
import { getExcelRowCount, getExcelHeaders, isExcelColumnPopulated } from 'excel-file-reader-ts';
// For JavaScript (CommonJS)
// const { getExcelRowCount, getExcelHeaders, isExcelColumnPopulated } = require('excel-file-reader-ts');

Let's assume you have an Excel file named data.xlsx in your project root with the following content (Sheet1):

# Create Or Use Sample excel file 

And another sheet named EmptySheet which is completely empty.

# Get Total Row Count
This function returns the total number of rows in a specified Excel sheet.

```javascript

import { getExcelRowCount } from 'excel-file-reader-ts';
import * as path from 'path';

const filePath = path.join(__dirname, 'data.xlsx'); // Adjust path as needed

async function countRows() {
  try {
    // Get row count for the first sheet (default)
    const totalRows = await getExcelRowCount(filePath);
    console.log(`Total rows in default sheet: ${totalRows}`); // Expected: 5 (1 header + 4 data)

    // If you have a specific sheet name
    // const sheet2Rows = await getExcelRowCount(filePath, 'AnotherSheet');
    // console.log(`Total rows in 'AnotherSheet': ${sheet2Rows}`);

    // Example for an empty sheet
    const emptySheetPath = path.join(__dirname, 'empty_data.xlsx'); // Assume this file has an empty sheet
    // You might need to create an empty_data.xlsx with an empty sheet named 'EmptySheet'
    // For demonstration, let's assume 'data.xlsx' has an empty sheet named 'EmptySheet'
    const emptyRows = await getExcelRowCount(filePath, 'EmptySheet');
    console.log(`Total rows in 'EmptySheet': ${emptyRows}`); // Expected: 0

  } catch (error: any) {
    console.error(`Error getting row count: ${error.message}`);
  }
}

countRows();
```

# Get Headers
This function retrieves an array of header names from the first row of an Excel sheet.
```javascript

import { getExcelHeaders } from 'excel-file-reader-ts';
import * as path from 'path';

const filePath = path.join(__dirname, 'data.xlsx'); // Adjust path as needed

async function fetchHeaders() {
  try {
    // Get headers from the first sheet (default)
    const headers = await getExcelHeaders(filePath);
    console.log('Headers:', headers); // Expected: ['Product ID', 'Product Name', 'Price', 'Stock', 'Description']

    // If you have a specific sheet name
    // const sheet2Headers = await getExcelHeaders(filePath, 'AnotherSheet');
    // console.log('Headers from AnotherSheet:', sheet2Headers);

  } catch (error: any) {
    console.error(`Error getting headers: ${error.message}`);
  }
}

fetchHeaders();

```
# Check if Column is Populated
This function checks if a column, identified by its header name, contains any non-empty values in its data rows. It returns true if at least one cell in that column (excluding the header) has a value, false otherwise.
```javascript
import { isExcelColumnPopulated } from 'excel-file-reader-ts';
import * as path from 'path';

const filePath = path.join(__dirname, 'data.xlsx'); // Adjust path as needed

async function checkColumnPopulation() {
  try {
    // Check a populated column (e.g., 'Product Name')
    const isProductNamePopulated = await isExcelColumnPopulated(filePath, 'Product Name');
    console.log(`Is 'Product Name' column populated? ${isProductNamePopulated}`); // Expected: true

    // Check a partially populated column (e.g., 'Description')
    const isDescriptionPopulated = await isExcelColumnPopulated(filePath, 'Description');
    console.log(`Is 'Description' column populated? ${isDescriptionPopulated}`); // Expected: true (because 'Ergonomic' exists)

    // Check a completely empty column (if you had one, e.g., 'Notes')
    // For this example, let's assume 'data.xlsx' had a column 'Notes' with no values
    // const isNotesPopulated = await isExcelColumnPopulated(filePath, 'Notes');
    // console.log(`Is 'Notes' column populated? ${isNotesPopulated}`); // Expected: false

    // Check a non-existent header (should throw an error)
    try {
      await isExcelColumnPopulated(filePath, 'NonExistentHeader');
    } catch (error: any) {
      console.error(`Error for non-existent header (expected): ${error.message}`);
    }

  } catch (error: any) {
    console.error(`An unexpected error occurred: ${error.message}`);
  }
}

checkColumnPopulation();
```

# API Reference
All functions are asynchronous and return Promises.

`getExcelRowCount(filePath: string, sheetName?: string): Promise<number>`
    filePath (string):The absolute or relative path to the Excel file.

    sheetName (string, optional): The name of the specific sheet to read. If omitted, the first sheet in the workbook will be used.

    Returns: Promise<number> - The total number of rows in the specified sheet. Includes header rows.

`getExcelHeaders(filePath: string, sheetName?: string): Promise<string[]>`
    filePath (string): The absolute or relative path to the Excel file.

    sheetName (string, optional): The name of the specific sheet to read. If omitted, the first sheet in the workbook will be used.

    Returns: Promise<string[]> - An array of strings representing the header names from the first row of the specified sheet. Returns an empty array if no headers are found or the sheet is empty.

`isExcelColumnPopulated(filePath: string, headerName: string, sheetName?: string): Promise<boolean>`
    filePath (string): The absolute or relative path to the Excel file.

    headerName (string): The exact name of the header column to check for values. This is case-sensitive.

    sheetName (string, optional): The name of the specific sheet to read. If omitted, the first sheet in the workbook will be used.

    Returns: Promise<boolean> - true if the specified column contains at least one non-empty (not null, undefined, or whitespace-only string) value in its data rows; false otherwise.

# Error Handling
All functions will throw an Error (and reject their Promises) in the following scenarios:

The filePath does not point to an existing file.

The sheetName provided does not exist in the Excel workbook.

For isExcelColumnPopulated, if the headerName does not exist in the specified sheet.

If the Excel file is corrupted or cannot be parsed by the xlsx library.

It is recommended to wrap calls to these functions in try...catch blocks or use .catch() with Promises to handle potential errors gracefully.

# Github Repository Link
[https://github.com/rajasekar-arch/excel-file-reader-ts]

# Contributing
Contributions are welcome! If you find a bug or have a feature request, please open an issue on the GitHub repository.

# License
This project is licensed under the MIT License - see the LICENSE file for details.