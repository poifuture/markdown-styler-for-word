const devCerts = require("office-addin-dev-certs")
const { CleanWebpackPlugin } = require("clean-webpack-plugin")
const CopyWebpackPlugin = require("copy-webpack-plugin")
const ExtractTextPlugin = require("extract-text-webpack-plugin")
const HtmlWebpackPlugin = require("html-webpack-plugin")
const TerserPlugin = require("terser-webpack-plugin")
const webpack = require("webpack")
const { BundleAnalyzerPlugin } = require("webpack-bundle-analyzer")

module.exports = async (env, options) => {
  const config = {
    devtool: "source-map",
    entry: {
      taskpane: ["react-hot-loader/patch", "./src/taskpane/index.tsx"],
      commands: "./src/commands/commands.ts",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /is-plain-obj.*\.js$/,
          use: {
            loader: "babel-loader",
            options: { presets: ["@babel/preset-env"] },
          },
        },
        {
          test: /\.tsx?$/,
          use: ["react-hot-loader/webpack", "ts-loader"],
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
          use: {
            loader: "file-loader",
            query: {
              name: "assets/[name].[ext]",
            },
          },
        },
      ],
    },
    optimization: {
      splitChunks: {
        chunks: "all",
      },
      ...(options.mode === "production" && {
        minimizer: [
          new TerserPlugin({
            parallel: true,
            cache: true,
            sourceMap: true,
            terserOptions: {
              compress: {
                global_defs: {
                  DEBUG: false,
                },
                pure_funcs: ["console.debug", "console.info"],
              },
            },
          }),
        ],
      }),
    },
    plugins: [
      // Note: Bundle Analyzer will open a server and block "yarn build" at the end.
      // Use only when needed
      ...(options.mode === "development" ? [new BundleAnalyzerPlugin()] : []),
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin([
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css",
        },
      ]),
      new ExtractTextPlugin("[name].[hash].css"),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin([
        {
          from: "./assets",
          ignore: ["*.scss"],
          to: "assets",
        },
      ]),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      https:
        options.https !== undefined
          ? options.https
          : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  }
  return config
}
