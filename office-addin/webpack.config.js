const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// mkcert-generated certificates (trusted by Windows/Chrome/Edge via local CA)
const CERTS_DIR = path.resolve(__dirname, "certs");
const httpsOptions = {
  key: fs.readFileSync(path.join(CERTS_DIR, "localhost-key.pem")),
  cert: fs.readFileSync(path.join(CERTS_DIR, "localhost.pem")),
};

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.tsx",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/index.html",
      filename: "taskpane.html",
      chunks: ["taskpane"],
    }),
  ],
  devServer: {
    port: 3000,
    server: {
      type: "https",
      options: httpsOptions,
    },
    headers: { "Access-Control-Allow-Origin": "*" },
    hot: true,
  },
};
