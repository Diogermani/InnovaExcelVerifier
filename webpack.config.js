const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  devtool: "source-map",
  entry: {
    taskpane: "./src/taskpane/taskpane.ts",
    launchevent: "./src/launchevent/launchevent.ts"
  },
  resolve: {
    extensions: [".ts", ".js"]
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: "ts-loader"
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"]
      }
    ]
  },
  // Mark xlsx as external so it is NOT bundled into launchevent.js.
  // It will be loaded from the SheetJS CDN instead (see commands.html).
  // This reduces launchevent.js from ~909 KB to ~20 KB.
  externals: {
    xlsx: "XLSX"
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane"]
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["launchevent"]
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "manifest.xml",
          to: "manifest.xml"
        },
        {
          from: "src/config/config.json",
          to: "config.json"
        },
        {
          from: "assets/*",
          to: "assets/[name][ext]",
          noErrorOnMissing: true
        }
      ]
    })
  ],
  output: {
    filename: "[name].js",
    path: path.resolve(__dirname, "dist"),
    clean: true
  }
};
