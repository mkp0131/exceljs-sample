import express from 'express';
import bodyParser from 'body-parser';
import morgan from 'morgan';
import fs from 'fs';

const app = express();
const logger = morgan('dev');

app.use(logger);

// form 전송시 사용!
// html form 을 해석해서 JS object 로 생성
app.use(bodyParser.urlencoded({ extended: false }));
// json 객체 해석
app.use(express.json());

// 사진, 동영상 static 폴더
app.use('/uploads', express.static('uploads'));
// css, js static 폴더
app.use('/static', express.static('assets'));

app.set('view engine', 'pug');
app.set('views', process.cwd() + '/src/views');

// 라우터
app.get('/', (req, res) => {
  res.render('base');
})

import path from 'path';
import ExcelJS from 'exceljs';
app.get('/excel', (req, res) => {
  // 엑셀 파일 생성
  const workbook = new ExcelJS.Workbook();
  
  // 시트 생성('시트이름')
  const worksheet = workbook.addWorksheet('first sheet');
  
  // 특정 셀에 값넣기
  worksheet.getCell('A1').value = '엑셀 특정 셀에 값 넣기';
  
  const fileName = "newSaveeee" + Date.now() + ".xlsx";
  const filePath = path.join(process.cwd(), fileName);
  workbook.xlsx
    .writeFile(filePath)
    .then(response => {
      console.log("file is written");
      console.log(filePath);
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
})

// 모두 안걸리는 것 404 페이지 처리
app.use('/*', (req, res) => {
  res.send('404');
});

const PORT = process.env.PORT || 4000;

app.listen(PORT, () => {
  console.log(`⭐ Start Server PORT: ${PORT}`);
});