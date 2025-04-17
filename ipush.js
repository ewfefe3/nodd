
/*const express = require('express');
const app = express();
const port = 3000;
const xlsx = require('xlsx');
const fs = require('fs');
const cors = require('cors');
app.use(cors());

 workbook = xlsx.readFile('data.xlsx');
 sheetName = workbook.SheetNames.find(name => name === 'Sheet1');
 sheet = workbook.Sheets[sheetName];
 data = xlsx.utils.sheet_to_json(sheet);*/

//ipush('Sheet1','nam::yunes|a|pnam::yacoubi|a|pd::13300k|a|type::@time|a|info::@date|a|id::no ealse|a|class::proumear');



 // Export the function
 module.exports = {
  ipush,newxlsx,nowdate,nowtime
};
/*
module.exports = function ipush2(x, y) {
  return ipush20(x,y) ;
};*/



function ipush(sheetName,a,xlsx,filename) {
//  const workbook = xlsx.readFile('data.xlsx');
 // sheetName = workbook.SheetNames.find(name => name === sheetName);
   sheet = workbook.Sheets[sheetName];
   data = xlsx.utils.sheet_to_json(sheet);
   vals = a.split('|a|');
   worksheet = workbook.Sheets[sheetName];
   imessage = '';
  deflt_row = [];
  Isnew = false;
  range = false;
  // Step 3: Get the range of the sheet .
  if( !workbook.Sheets[sheetName])  {
  //console.log(` worksheet ${sheetName} not found !!!!!`);
  workbook = xlsx.readFile(filename);
  const data = [

  ];
  const newSheet = xlsx.utils.aoa_to_sheet(data);
  xlsx.utils.book_append_sheet(workbook, newSheet, sheetName);
  //xlsx.writeFile(workbook, 'data.xlsx');
  }
  //else console.log(` worksheet ${sheetName} exist`);

   if(worksheet != null) range = worksheet['!ref']; // This gives the sheet's range (e.g., 'A1:C10')
if (!range)  range = "A1:B1";
  // Step 4: Extract column letter (e.g., 'A', 'B', 'C')
  const columnLetter = 'A';  // Specify the column letter you are interested in (e.g., 'A' for the first column)

  // Step 5: Determine the last row of the sheet by checking the row numbers in the range
  lastRow = parseInt(range.split(':')[1].replace(/[A-Z]/g, ''), 10);

  // Step 6: Get the last cell in the specified column
  if(worksheet != null) lastCell = worksheet[columnLetter + lastRow]; // Access the last cell in the column
  else lastCell = 1;
 parsed = 1;
  if (lastCell) {
  Isnew = true;
  lastRow = parseInt(range.split(':')[1].replace(/[A-Z]/g, ''), 10);
  if(worksheet != null) lastCell = worksheet[columnLetter + lastRow];
   else lastCell = 1;
  parsed = parseInt(lastCell.v, 0)+1;
  //  console.log(`Last cell in column ${columnLetter}:`, lastCell.v); // Logs the value of the last cell
  } else {
    console.log(`No data found in column ${columnLetter}.`);
  }
  //--------------------------------------------------------------------------------------
  const last = isNaN(parsed) ? 0 : parsed;
  const data2 = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); // '...header: 1' gives a 2D array
  Frow = data2[0];
  let fruits = vals;
  const rowMap = new Map();
  var NULLS = [];
  var divaic = '';
for (let i = 0; i < fruits.length; i++) {
const fr = fruits[i].split("::");
if(fr[1] == '') NULLS.push(fr[0]);
rowMap.set(fr[0], fr[1]);
deflt_row.push(fr[0]);
if(i == 0) divaic = fr[1];
//console.log(fr);
}
if(NULLS.length == 0){
if( Frow == null) {
console.log(Isnew);
console.log(Frow);
deflt_row.unshift('ID');
data2.push(deflt_row);
Frow = deflt_row;
}
for (let i = 0; i < fruits.length; i++) {
fruits[i] = rowMap.get(Frow[i+1]);
}
  const replacements2 = {
    '@time': nowtime(),
    '@date': nowdate()
  };

  fruits = fruits.map(fruit => replacements2[fruit] || fruit);
  fruits.unshift(last.toString());
   data2.push(fruits);
  const newWorksheet = xlsx.utils.aoa_to_sheet(data2); // 'aoa_to_sheet' converts 2D array to sheet
  workbook.Sheets[sheetName] = newWorksheet; // Update the first sheet
  xlsx.writeFile(workbook, filename);
    imessage = `now put [${a}] in (${sheetName}) is Ok`;
  }else imessage = ` error::100::input ${NULLS} not completed`;

  return imessage;
}
function nowtime(){
  const now = new Date();
const hours = now.getHours();
const minutes = now.getMinutes();
const seconds = now.getSeconds();
return `${hours}:${minutes}:${seconds}`;
}
function nowdate(){
  const now = new Date();
  const year = now.getFullYear();
  const month = (now.getMonth() + 1).toString().padStart(2, '0'); // getMonth() returns 0-11, so add 1
  const day = now.getDate().toString().padStart(2, '0');
  return `${year}-${month}-${day}`;
}
function newxlsx({ isheet = 'Sheet1' } = {},{ filePath = 'new-file' } = {}) {
const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');
 // const filePath = 'new-file.xlsx';
  filePath = filePath+'.xlsx';
 dir = filePath.split('/')[0];
 //filePath = filePath.split('/')[1];
  const filePath2 = path.join(dir, 'new-file.xlsx');

  // Ensure the directory exists
  if (!fs.existsSync(dir)) {
    console.log(`Directory "${dir}" does not exist. Creating it...`);
    fs.mkdirSync(dir, { recursive: true });
  }else dir = '';
  // Check if the file exists
  if (fs.existsSync(filePath)) {
 //   console.log(`The file "${filePath}" already exists.`);
  } else {
    // Create a new workbook
    const wb = XLSX.utils.book_new();

    // Create some data to insert into the Excel sheet (can add data here if necessary)
    const data = [];  // Empty data for now

    // Convert the data into a worksheet
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Append the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, isheet);

    // Write the workbook to a new .xlsx file

    XLSX.writeFile(wb, filePath);

    console.log(`Excel file "${filePath}" created successfully!`);
  }
}



