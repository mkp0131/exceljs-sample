// # Sever에서 파일 내려온것 다운로드
// fetch('/excel').then( res => res.blob() )
// .then( blob => {
//   var file = window.URL.createObjectURL(blob);
//   const url = window.URL.createObjectURL(blob);
//   const anchor = document.createElement('a');
//   anchor.href = url;
//   // 파일명
//   anchor.download = `테스트.xlsx`;
//   anchor.click();
//   window.URL.revokeObjectURL(url);
// });

// # 브라우져 사이드에서 사용
function makeExcel(data) {
  // 엑셀 파일 생성
  const workbook = new ExcelJS.Workbook();

  // 시트 생성('시트이름')
  const worksheet1 = workbook.addWorksheet('일반');

  // 특정 셀에 값넣기
  // worksheet1.getCell('A1').value = '엑셀 특정 셀에 값 넣기';

  const fileName = "newSaveeee" + Date.now() + ".xlsx";

  // Header 설정
  const columns = Object.keys(data[0]).map((item) => {
    return {
      header: item,
      key: item,
      width: 40,
      style: { font: { size: 14 }, numFmt: '@', alignment: { horizontal: 'center' } },
    }
  });

  worksheet1.columns = columns;

  // Rows 넣기
  worksheet1.insertRows(2, data);


  // Header 부분 스타일 재지정
  // worksheet1.columns 의 스타일을 따라가기 때문에 스타일을 재지정해준다.
  const headerStyle = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'cce6ff' },
  };

  const headerBorderStyle = {
    left: { style: 'thin', color: { argb: 'bfbfbf' } },
    right: { style: 'thin', color: { argb: 'bfbfbf' } },
  };

  for (let i = 1; i <= worksheet1.columnCount; i++) {
    const headerEachCell = worksheet1.getCell(`${String.fromCharCode(i + 64)}1`);
    headerEachCell.fill = headerStyle;
    headerEachCell.border = headerBorderStyle;
    headerEachCell.alignment = { horizontal: 'center' };
  }

  // 생성한 데이터의 위치를 옮긴다.
  // 1번과 0번 rows 사이에 빈칸을 3개 삽입
  worksheet1.spliceRows(1, 0, [], []);

  // 새로생긴 빈칸에 제목 텍스트를 적어준다.
  worksheet1.getCell('A1').value = {
    richText: [
      { 
        text: '일반 엑셀입니다. 넘쳤을때 테스트 입니다.', 
        font: {
          size: 15,
          color: {
            argb: 'A52A2A'
          },
          bold: true,
        },
      },
    ],
    alignment: {
      horizontal: 'left'
    }
  };
  worksheet1.getCell('A1').alignment = { horizontal: 'left' };

  // 테이블 추가
  const worksheet2 = workbook.addWorksheet('테이블');

  const tableData = data.map(item => Object.values(item));

  console.log('@', tableData);

  worksheet2.addTable({
    name: 'letsMakeTable',
    ref: 'A3', // 시작할 Cell
    headerRow: true,
    totalsRow: false, // 맨 아래 토탈 생성
    style: {
      theme: 'TableStyleLight1',
      showRowStripes: true,
    },
    columns: [
      { name: 'id', filterButton: true },
      { name: 'name', filterButton: true },
      { name: 'D.O.B', filterButton: true },
      { name: 'salary', filterButton: true, },
    ],
    rows: tableData,
  });


  worksheet2.eachRow((row, rowNo) => {
    row.height = 28;
    row.alignment = {
      vertical: 'middle',
      horizontal: 'center'
    }
    row.eachCell((cell, colNo) => {
      if (cell.value || cell.value === 0) {
        const eachCell = row.getCell(colNo);
        eachCell.font = { size: 14 };

        if (colNo === 6) eachCell.numFmt = '$#,##0';
      }
    });
  });

  for (let i = 1; i <= worksheet2.columnCount; i++) {
    const eachColumn = worksheet2.getColumn(i);

    // if (i === 2) {
    //   eachColumn.alignment = { horizontal: 'right' };
    // } else eachColumn.alignment = { horizontal: 'center' };

    if (eachColumn.values.length !== 0) {
      eachColumn.width = 20;
    }
  }

  // 새로생긴 빈칸에 제목 텍스트를 적어준다.
  worksheet2.getCell('A1').value = {
    richText: [
      { 
        text: '테이블 제목', 
        font: {
          size: 15,
          color: {
            argb: 'A52A2A'
          },
          bold: true,
        },
      },
    ],
    alignment: {
      horizontal: 'left'
    }
  };
  worksheet2.getCell('A1').alignment = { horizontal: 'left' };

  // 생성한 엑셀파일 다운로드
  workbook.xlsx.writeBuffer().then((data) => {
    // return;
    const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    // 파일명
    anchor.download = `브라우져.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  })
}


fetch('/json')
  .then(data => data.json())
  .then(json => {
    makeExcel(json);
  });



