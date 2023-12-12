const XlsxPopulate = require("xlsx-populate");
const readXlsxFile = require("read-excel-file/node");

const cors = require("cors");

const express = require("express");
const app = express();
const port = 3000;
app.use(cors());

const readFromexcel = () => {
  readXlsxFile("result.xlsx").then((rows) => {    
    app.get("/", (req, res) => {
      res.send(rows);
    });
  });
};

const addToExcel = () => {
  XlsxPopulate.fromBlankAsync().then((workbook) => {
    workbook.sheet("Sheet1").cell("A1").value("Name");
    workbook.sheet("Sheet1").cell("B1").value("Age");
    workbook.sheet("Sheet1").cell("C1").value("Address");
    workbook.sheet("Sheet1").cell("D1").value("Phone");
    workbook.sheet("Sheet1").cell("A2").value("Pierre");
    workbook.sheet("Sheet1").cell("B2").value("27");
    workbook.sheet("Sheet1").cell("C2").value("Zgharta");
    workbook.sheet("Sheet1").cell("D2").value("70913457");
    workbook.sheet("Sheet1").cell("A3").value("Jawad");
    workbook.sheet("Sheet1").cell("B3").value("23");
    workbook.sheet("Sheet1").cell("C3").value("Zgharta");
    workbook.sheet("Sheet1").cell("D3").value("78291831");
    workbook.sheet("Sheet1").cell("A4").value("Saliba");
    workbook.sheet("Sheet1").cell("B4").value("26");
    workbook.sheet("Sheet1").cell("C4").value("Zgharta");
    workbook.sheet("Sheet1").cell("D4").value("76435216");


    return workbook.toFileAsync("result.xlsx");
  });
};

readFromexcel();

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});

// readFromexcel()
addToExcel()
