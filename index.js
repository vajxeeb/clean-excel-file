
//Import library
let XLSX = require('xlsx');
//Declare workbook for read file excel
let workbook = XLSX.readFile('before-file.xlsx');
//Declare WSH for get Sheet name or index of sheet in excel workbook of excel
let WSH = workbook.SheetNames[0];
let worksheet = workbook.Sheets[WSH];

 // Declare valiable for get the values in each cell in excel file  
const level = [];
const columnA_Id = [];
const columnB_Sn1 = [];
const columnC_Sn2 = [];
const columnD_Sn3 = [];
const columnE_Sn4 = [];
const columnF_Sn5 = [];
const columnG_Sn6 = [];


//Push each values in cell  to each valiable that we delare above
for (let z in worksheet) {

  if (z.toString()[0] === 'A') {
    level.push(worksheet[z].v);
  }
  if (z.toString()[0] === 'B') {
    columnA_Id.push(worksheet[z].v);
  }
  if (z.toString()[0] === 'C') {
    columnB_Sn1.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'D') {
    columnC_Sn2.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'E') {
    columnD_Sn3.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'F') {
    columnE_Sn4.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'G') {
    columnF_Sn5.push(worksheet[z].v);

  }
  if (z.toString()[0] === 'H') {
    columnG_Sn6.push(worksheet[z].v);

  }

}
//Delete 2 row of each array like below
level.splice(0, 2)
columnA_Id.splice(0, 2)
columnB_Sn1.splice(0, 2)
columnC_Sn2.splice(0, 2)
columnD_Sn3.splice(0, 2)
columnE_Sn4.splice(0, 2)
columnF_Sn5.splice(0, 2)
columnG_Sn6.splice(0, 2)

//Prepare for update the file back
const Excel = require('exceljs');
const WB = new Excel.Workbook();

WB.xlsx.readFile('before-file.xlsx')
  .then(function () {

    const worksheet = WB.getWorksheet(1);
    let row1 = worksheet.getRow(1);


    //Delete the vlues of 2 row above
    let clcell
    for (let i = 1; i <= 10; i++) {

      clcell = worksheet.getRow(1)
      clcell.getCell(i).value = ""

      clcell = worksheet.getRow(2)
      clcell.getCell(i).value = ""

    }
    //.....Set title with the first row.......
    row1.getCell(1).value = "ລຳດັບ"
    row1.getCell(2).value = "ລະຫັດເລກ"
    row1.getCell(3).value = "ເລກ 1 ໂຕ"
    row1.getCell(4).value = "ເລກ 2 ໂຕ"
    row1.getCell(5).value = "ເລກ 3 ໂຕ"
    row1.getCell(6).value = "ເລກ 4 ໂຕ"
    row1.getCell(7).value = "ເລກ 5 ໂຕ"
    row1.getCell(8).value = "ເລກ 6 ໂຕ"

    //.....Set  values to each cell .......
    let rowindex
    for (let cell = 2; cell <= columnA_Id.length; cell++) {

      rowindex = worksheet.getRow(cell)
      //Set values with each cell
      rowindex.getCell(1).value = level[cell-2]
      rowindex.getCell(2).value = columnA_Id[cell-2]
      rowindex.getCell(3).value = columnB_Sn1[cell-2]
      rowindex.getCell(4).value = columnC_Sn2[cell-2]
      rowindex.getCell(5).value = columnD_Sn3[cell-2]
      rowindex.getCell(6).value = columnE_Sn4[cell-2]
      rowindex.getCell(7).value = columnF_Sn5[cell-2]
      rowindex.getCell(8).value = columnG_Sn6[cell-2]
    }

    row1.commit();
    rowindex.commit();
    clcell.commit();

    //Beging writting back a file 
    return WB.xlsx.writeFile('After-file.xlsx');
  })



