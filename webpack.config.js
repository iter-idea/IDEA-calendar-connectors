var nodeExternals = require('webpack-node-externals');

module.exports = {
  name: 'prod',
  entry: './src/index.ts',
  resolve: { extensions: ['.ts', '.tsx', '.js'] },
  module: { rules: [{ test: /\.tsx?$/, loader: 'ts-loader', options: { transpileOnly: false } }] },
  output: {
    filename: 'index.js',
    /* path: SPECIFIED BY COMMAND LINE */
    libraryTarget: 'umd'
  },
  target: 'node',
  externals: [nodeExternals()],
  mode: 'production',
  optimization: { minimize: false },
  stats: 'normal'
};
