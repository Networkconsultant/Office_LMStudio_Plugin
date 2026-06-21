const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");

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
    // HTTP only — AllowHTTP=1 is set in the Office registry (HKCU\...\WEF\AllowHTTP)
    headers: { "Access-Control-Allow-Origin": "*" },
    hot: true,
    // Serve taskpane.html for / and any unmatched route
    devMiddleware: {
      index: "taskpane.html",
    },
    historyApiFallback: {
      index: "/taskpane.html",
      rewrites: [
        // Don't rewrite /test — let setupMiddlewares handle it
        { from: /^\/test/, to: "/test" },
      ],
    },
    setupMiddlewares(middlewares, devServer) {
      // Simple diagnostic page — confirm WebView2 can reach localhost
      devServer.app.get("/test", (_req, res) => {
        res.send(`<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>LMStudio AI - Test</title></head>
<body style="font-family:sans-serif;padding:20px;background:#1a1a2e;color:#eee">
<h1 style="color:#4CAF50">Server is reachable!</h1>
<p>If you see this in the Office task pane, the connection is working.</p>
<p>Time: ${new Date().toISOString()}</p>
<p><a href="/taskpane.html" style="color:#4fc3f7">Open full add-in</a></p>
</body></html>`);
      });
      return middlewares;
    },
  },
};
