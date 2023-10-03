import fs from 'fs';
import xml2js from 'xml2js';
import excel4node from 'excel4node';

// Function to read XML content from a file
function readXml(filePath) {
  try {
    const xmlData = fs.readFileSync(filePath, 'utf-8');
    return xmlData;
  } catch (error) {
    console.error('Error reading XML file:', error);
    throw error;
  }
}

// Function to parse XML to Javascript Object
function parseXmlToJSON(xmlData, elementName) {
  try {
    return new Promise((resolve, reject) => {
      const parser = new xml2js.Parser({ explicitArray: false });
      parser.parseString(xmlData, (err, result) => {
        if (err) {
          console.error('Error parsing XML to JSON:', err);
          reject(err);
        } else {
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
function convertJsonToExcel(jsonData, fileName) {
  const wb = new excel4node.Workbook();
  const ws = wb.addWorksheet('employee_details');

  const headerStyle = wb.createStyle({
    font: {
      bold: true,
      color: '#000000',
    },
    fill: {
      type: 'pattern',
      patternType: 'solid',
      fgColor: '#F4E869',
    },
    border: {
      left: { style: 'thin', color: 'black' },
      right: { style: 'thin', color: 'black' },
      top: { style: 'thin', color: 'black' },
      bottom: { style: 'thin', color: 'black' },
    },
    alignment: {
      horizontal: 'center',
    },
  });

  const dataStyle = wb.createStyle({
    font: {
      color: '#000000',
    },
    border: {
      left: { style: 'thin', color: 'black' },
      right: { style: 'thin', color: 'black' },
      top: { style: 'thin', color: 'black' },
      bottom: { style: 'thin', color: 'black' },
    },
    alignment: {
      horizontal: 'center',
    },
  });

  const headingColumnNames = Object.keys(jsonData[0]);
  let headingColumnIndex = 1;
  headingColumnNames.forEach((heading) => {
    ws.cell(2, headingColumnIndex++)
      .string(heading.charAt(0).toUpperCase() + heading.slice(1))
      .style(headerStyle);
  });

  let rowIndex = 3;
  jsonData.forEach((record) => {
    let columnIndex = 1;
    Object.keys(record).forEach((columnName) => {
      ws.cell(rowIndex, columnIndex++)
        .string(record[columnName])
        .style(dataStyle);
    });
    rowIndex++;
  });

  headingColumnIndex = 1;
  Object.keys(jsonData[0]).forEach(() => {
    ws.column(headingColumnIndex++).setWidth(20); // Adjust the width as needed
  });

  wb.write(fileName);
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
const xmlFilePath = 'employee.xml';
const excelFileName = 'employee_details.xlsx';
const elementName = 'employee';
processXmlFile(xmlFilePath, excelFileName, elementName);
