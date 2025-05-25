const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  entry: {
    taskpane: './src/taskpane/taskpane.ts',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].bundle.js',
    clean: true
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.html', '.js']
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: 'ts-loader'
      },
      {
        test: /\.html$/,
        exclude: /node_modules/,
        use: 'html-loader'
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: 'taskpane.html',
      template: './src/taskpane/taskpane.html',
      chunks: ['taskpane']
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: 'assets/*',
          to: 'assets/[name][ext]'
        },
        {
          from: 'manifest.xml',
          to: 'manifest.xml'
        }
      ]
    })
  ],
  devServer: {
    headers: {
      "Access-Control-Allow-Origin": "*"
    },
    https: true,
    port: 3000
  }
}; 