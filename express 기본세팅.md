# express 기본세팅

## 패키지 설치

```sh
npm i @babel/cli @babel/core @babel/node @babel/preset-env babel-loader nodemon webpack webpack-cli --save-dev

npm i body-parser dotenv express express-session morgan node-fetch pug regenerator-runtime bcrypt
```

## package.json

```json
"scripts": {
  "build:server": "babel src -d build",
  "dev:server": "nodemon",
},
```

## nodemon 세팅

- `nodemon.json` 파일 생성

```json
{
  "ignore": ["src/client/*", "assets/*", "webpack.config.js"],
  "exec": "babel-node src/server.js"
}
```

## /src/server.js 세팅

```js
import express from 'express';
import bodyParser from 'body-parser';
import morgan from 'morgan';

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

// 모두 안걸리는 것 404 페이지 처리
app.use('/*', (req, res) => {
  res.send('404');
});

const PORT = process.env.PORT || 4000;

app.listen(PORT, () => {
  console.log(`⭐ Start Server PORT: ${PORT}`);
});
```

