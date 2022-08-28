var path = require('path');

module.exports = {
  entry: './src/index.js',
  target: ['web', 'es5'], // IE11 대응
  output: {
    path: path.resolve(__dirname, 'build'),
    filename: 'main.bundle.js',
  },
  module: {
    rules: [
      {
        test: /\.m?js$/,
        exclude: /(node_modules|bower_components)/,
        use: {
          loader: 'babel-loader',
          options: {
            presets: ['@babel/preset-env'],
          },
        },
      },
    ],
  },
  // devtool: 'source-map', // 개발에서 소스코드 보고싶을때 사용
};