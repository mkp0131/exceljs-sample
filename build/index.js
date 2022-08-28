"use strict";

var _exceljs = _interopRequireDefault(require("exceljs"));

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }

// 엑셀 파일 생성
var workbook = new _exceljs["default"].Workbook();
workbook.xlsx.readFile(process.env.INIT_CWD + '/xlsx/sample.xlsx').then(function () {
  var worksheet = workbook.getWorksheet('first sheet');
  worksheet.eachRow({
    includeEmpty: true
  }, function (row, rowNumber) {
    console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
  });
});