// webpack.config.js
/* eslint-disable @typescript-eslint/no-var-requires */
const path = require('path');
const ESLintPlugin = require('eslint-webpack-plugin');
const GasPlugin = require('gas-webpack-plugin');

module.exports = {
  mode: 'development',
  devtool: false,
  context: __dirname,
  entry: {
    main: './src/main.ts',
  },
  output: {
    path: path.join(__dirname, 'dist'),
    filename: '[name].js',
  },
  resolve: {
    extensions: ['.ts', '.js'],
  },
  module: {
    rules: [
      {
        test: /\.[tj]s$/,
        exclude: /node_modules/,
        loader: 'babel-loader',
      },
    ],
  },
  plugins: [new ESLintPlugin({ exclude: 'node_modules' }), new GasPlugin()],
};