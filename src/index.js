import path from 'path';
import ExcelJS from 'exceljs';

// 엑셀 파일 생성
const workbook = new ExcelJS.Workbook();

// 시트 생성('시트이름')
const worksheet = workbook.addWorksheet('first sheet');

// 특정 셀에 값넣기
worksheet.getCell('A1').value = '엑셀 특정 셀에 값 넣기';

// worksheet.addRow([3, 'Michael']).commit();

const fileName = "newSaveeee" + Date.now() + ".xlsx";
const filePath = path.join(process.cwd(), fileName);
workbook.xlsx
  .writeFile(filePath)
  .then(response => {
    console.log("file is written");
    console.log(filePath);
    res.sendFile(filePath);
  })
  .catch(err => {
    console.log(err);
  });

