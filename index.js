// Import necessary libraries
import fs from 'fs';
import xml2js from 'xml2js';
import excel4node from 'excel4node';

// Function to read XML content from a file
const readXml = (filePath) => {
  try {
    const xmlData = fs.readFileSync(filePath, 'utf-8');
    return xmlData;
  } catch (error) {
    console.error('Error reading XML file:', error);
    throw error;
  }
}

// Function to parse XML to Javascript Object
const parseXmlToJSON = (xmlData, elementName) => {
  try {
    return new Promise((resolve, reject) => {
      // Use xml2js library to parse XML content
      const parser = new xml2js.Parser({ explicitArray: false });
      parser.parseString(xmlData, (err, result) => {
        if (err) {
          console.error('Error parsing XML to JSON:', err);
          reject(err);
        } else {
          // Extract the specified element from the parsed XML
          const tagName = Object.keys(result)[0];
          const jsonData = result[tagName][elementName];
          resolve(jsonData);
        }
      });
    });
  } catch (error) {
    console.error(error);
    throw error;
  }
}

// Function to convert Javascript Object to Excel
const convertJsonToExcel = (jsonData, fileName) => {
  // Create a new Excel workbook and worksheet
  const workBook = new excel4node.Workbook();
  const workSheet = workBook.addWorksheet('employee_details');

  // Define styles for header and data cells
  const headerStyle = workBook.createStyle({ /* ... */ });
  const dataStyle = workBook.createStyle({ /* ... */ });

  // Write header row with specified styles
  const headingColumnNames = Object.keys(jsonData[0]);
  let headingColumnIndex = 1;
  headingColumnNames.forEach((heading) => {
    workSheet.cell(2, headingColumnIndex++)
      .string(heading.charAt(0).toUpperCase() + heading.slice(1))
      .style(headerStyle);
  });

  // Write data rows with specified styles
  let rowIndex = 3;
  jsonData.forEach((record) => {
    let columnIndex = 1;
    Object.keys(record).forEach((columnName) => {
      workSheet.cell(rowIndex, columnIndex++)
        .string(record[columnName])
        .style(dataStyle);
    });
    rowIndex++;
  });

  // Set column widths and write the workbook to a file
  headingColumnIndex = 1;
  Object.keys(jsonData[0]).forEach(() => {
    workSheet.column(headingColumnIndex++).setWidth(20); // Adjust the width as needed
  });

  workBook.write(fileName);
  console.log('File created successfully.');
}

// Main function to execute the entire process
async function processXmlFile(filePath, excelFileName, elementName) {
  try {
    // Read XML file content
    const xmlData = readXml(filePath);

    // Parse XML to JSON
    const jsonData = await parseXmlToJSON(xmlData, elementName);

    // Convert JSON to Excel
    convertJsonToExcel(jsonData, excelFileName);
  } catch (error) {
    console.error('Error:', error);
  }
}

// Usage:
// Specify XML file path, Excel file name, and the XML element to extract
const xmlFilePath = 'employee.xml';
const excelFileName = 'employee_details.xlsx';
const elementName = 'employee';

// Execute the main processing function
processXmlFile(xmlFilePath, excelFileName, elementName);
