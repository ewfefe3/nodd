//npm install express

//npm install cors


// npm install xlsx

const express = require('express');
const app = express();
const port = 3000;
const xlsx = require('xlsx');
const fs = require('fs');
const cors = require('cors');
app.use(cors());
message = '';

 workbook = xlsx.readFile('data.xlsx');
 sheetName = workbook.SheetNames.find(name => name === 'Sheet1');
 sheet = workbook.Sheets[sheetName];
 data = xlsx.utils.sheet_to_json(sheet);
 function intalize(file,sheet){
 workbook = xlsx.readFile(file);
 sheetName = workbook.SheetNames.find(name => name === sheet);
 sheet = workbook.Sheets[sheetName];
 data = xlsx.utils.sheet_to_json(sheet);
}
const otherFile = require('./ipush');
const GoogleSheet = require('./GoogleSheet');
//otherFile.someFunction();
//GoogleSheet.ipussh();

app.get('/ENTER', (req, res) => {
  const a = req.query.a;
console.log(`ENTER: ${a} at : ${otherFile.nowtime()}`);
  res.json({ message: `OK` });
});
app.get('/GET', (req, res) => {
  const a = req.query.a;
  const b = req.query.b;
  const user = req.query.user;
  const Epass = req.query.Epass;
  const pass = req.query.pass;
  const sh = req.query.sh;

  if (!a || !b) {
    return res.status(400).send('enter a and b');
  }
 // console.log(` a :${a} b :${b}   user :${user}    sh :${sh}    pass :${pass}`);
  data = xlsx.utils.sheet_to_json(workbook.Sheets[sh]);
//**const filteredData = data.filter(row => row["nam"] === a && row["pnam"] === 'yacoubi');
const filteredData = data.filter(row => row[user] === a/*&& row[Epass] === 'pass'*/);
//console.log(`filteredData: ${JSON.stringify(filteredData, null, 2)}`);
if(filteredData.length > 0)
  res.json({ filteredData: filteredData });
  else res.json({ message: `error::101::${user} : ${a} are not Exist or passowrd false` });
});

app.get('/NEWW',async (req, res) => {
   a = req.query.a;
   b = req.query.b;
   f = req.query.force;
  if (!a|| !b) {
    return res.status(400).send('enter a and b.');
  }

   intalize('data.xlsx',b);
   sheetName = workbook.SheetNames.find(name => name === b);
    newA = a.replace(/%7C/g, "|");  // Replaces first "o" with "a"
    a = newA;
     sheet = workbook.Sheets[b];
     const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
     const column = jsonData.map(row => row[1]);
    a2 = a.split('|a|');
    marg3 = a2[0].split('::')[0];
    //console.log(` marg3 :${marg3} `);
    exist = false; user = '';
    for (let i = 0; i < a2.length; i++) {
    a3 = a2[i].split('::');
    if(a3[0] == marg3 && column.includes(a3[1])) {exist = true; user = a3[1];}
    }
    if(exist == true && f != 'true') {
    console.log(`${user} are Exist`);
    message = `error::101::${user} are Exist`;
    }
    else {
    message = otherFile.ipush(b,a,xlsx,'data.xlsx');
    message =await GoogleSheet.Gpush({
    sheet: 'Users 02',
    a: a,
    forc: true,
    });
    }
    otherFile.newxlsx({ isheet: 'G' },{ filePath: 'ecel/filePath' });
//res.json({ message: `now put [${a}] is Ok` });
res.json({ message: message });
});
//ipush('Sheet1','nam::yunes|a|pnam::yacoubi|a|pd::13300k|a|type::@time|a|info::@date|a|id::no ealse|a|class::proumear');












app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});


