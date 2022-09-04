# exceljs 사용법

- github 저장소: https://github.com/mkp0131/exceljs-sample

## 세팅

- cdn 모음: https://cdnjs.com/libraries/exceljs
- 바벨 폴리필 필수!
- 파일을 읽는 것은 NODE 환경에서만 가능!

## 기본사용법

### 브라우저 환경

```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/babel-polyfill/6.26.0/polyfill.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js" integrity="sha512-UnrKxsCMN9hFk7M56t4I4ckB4N/2HHi0w/7+B/1JsXIX3DmyBcsGpT3/BsuZMZf+6mAr0vP81syWtfynHJ69JA==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
```

```js
// 엑셀 파일 생성
const workbook = new ExcelJS.Workbook();

// 시트 생성('시트이름')
const worksheet = workbook.addWorksheet('first sheet');

// 특정 셀에 값넣기
worksheet.getCell('A1').value = '엑셀 특정 셀에 값 넣기';

// 생성한 엑셀파일 다운로드
workbook.xlsx.writeBuffer().then((data) => {
  const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  // 파일명
  anchor.download = `테스트.xlsx`;
  anchor.click();
  window.URL.revokeObjectURL(url);
})
```

### NODE 환경

- 파일읽기

```js
import ExcelJS from 'exceljs';

// 엑셀 파일 생성
const workbook = new ExcelJS.Workbook();

// 엑셀 파일 읽기
workbook.xlsx.readFile('../xlsx/sample.xlsx').then(function() {
  var worksheet = workbook.getWorksheet('first sheet');
  worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
  });
});
```

- 파일쓰기

```js
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
    // express 를 활용하여 파일 내보내기
    res.sendFile(filePath, err => {
      if(err) {
        console.log('파일정송 에러발생')
      }
      // 전송후 파일 삭제
      fs.unlink(filePath, (err) => {
        // log any error
        if (err) {
          console.log(err);
        }
      });
    });
  })
  .catch(err => {
    console.log(err);
  });
```

- ajax 로 쓴파일 받기

```js
fetch('/excel').then( res => res.blob() )
.then( blob => {
  var file = window.URL.createObjectURL(blob);
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  // 파일명
  anchor.download = `테스트.xlsx`;
  anchor.click();
  window.URL.revokeObjectURL(url);
});
```

### 📌 사용팁!

#### CSS 추가하여 값 입력하기!

```js
worksheet.getCell('A2').value = {
  richText: [{ text: 'css 추가합니다요',
  font: { size: 15, color: { argb: 'A52A2A' } } }],
};
```

#### Table 넣기

##### Header 세팅하기!

- columns는 엑셀의 필드값을 지정해준다. header에 필드명, key값은 해당 Header의 키값으로 나중에 데이터를 넣을 때 이 키를 기준으로 데이터가 쏙쏙 맞춰서 들어가기 때문에 중요

```js
worksheet.columns = [
  { header: 'Id', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 32 },
  { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 },
];
```

##### insert Rows 

```js
const data = [
  {
    id: 1,
    name: 'Jamey',
    DOB: '2022-12-25',
  },
  {
    DOB: '2100-01-10',
    name: 'Jimmy',
    id: 2,
  },
  {
    id: 3,
    name: 'Jesus',
    DOB: '2000-12-25',
  },
];

// 한줄만 넣고 싶을 경우
worksheet.insertRow(2, { id: 0, name: 'Jenny', DOB: '2020-11-11' });
// 많은 줄을 넣고 싶을 경우 Rows를 사용한다.
worksheet.insertRows(3, data);
```

