// Copyright (c) Wictor Wilén. All rights reserved. 
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

const webpack = require("webpack");
const nodeExternals = require("webpack-node-externals");
// const ESLintPlugin = require("eslint-webpack-plugin");
// const ForkTsCheckerWebpackPlugin = require("fork-ts-checker-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const NodemonPlugin = require('nodemon-webpack-plugin');

const path = require("path");
const fs = require("fs");
const argv = require("yargs").argv;

const debug = argv.debug !== undefined;
const lint = !(argv["no-linting"] || argv.l === true);
const BundleAnalyzerPlugin =
    require("webpack-bundle-analyzer").BundleAnalyzerPlugin;


const config = [
    {
        entry: {
            vendor: [
                'react',
                'react-dom',
                '@fluentui/react'
            ],
            taskpane: {
                import: [path.join(__dirname, "/src/client/taskpane/index.tsx")],
                dependOn: "vendor"
            }
            // commands: {
            //     import: [path.join(__dirname, "/src/client/commands/commands.ts")],
            //     dependOn: "vendor"
            // }
            // taskpane: [
            //     //path.join(__dirname, "/src/client/client.tsx"),
            //     path.join(__dirname, "/src/client/taskpane/index.tsx")
            //     //path.join(__dirname, "/src/client/commands/commands.tsx")
            // ],
            // commands: [
            //     //path.join(__dirname, "/src/client/client.tsx"),
            //     path.join(__dirname, "/src/client/commands/commands.ts")
            // ]
        },
        mode: debug ? "development" : "production",
        output: {
            path: path.join(__dirname, "/dist/web/scripts"),
            filename: "[name].js",
            libraryTarget: "umd",
            library: "cde-helper",
            publicPath: "scripts/" // relative to the HTML file
        },
        externals: {},
        devtool: "eval-source-map", //"source-map",
        resolve: {
            extensions: [".ts", ".tsx", ".js"],
            alias: {}
        },
        //target: "web",
        module: {
            rules: [{
                test: /\.tsx?$/,
                exclude: /node_modules/,
                use: {
                    loader: "ts-loader",
                    options: {
                        transpileOnly: true
                    }
                }
            },
            {
                test: /\.css$/,
                use: ['style-loader', 'css-loader']
            }],
            noParse: /node_modules\/quill\/dist/   //avoid react-quill warning info
        },
        plugins: [
            //new webpack.EnvironmentPlugin({ PUBLIC_HOSTNAME: undefined, TAB_APP_ID: null, TAB_APP_URI: null }),
            // new ForkTsCheckerWebpackPlugin({
            //     typescript: {
            //         configFile: "./src/client/tsconfig.json"
            //     }
            // }),
            new HtmlWebpackPlugin(
                {
                    filename: "../taskpane.html", // the dest folder is based on dist/web/scripts
                    template: "./src/public/taskpane.html",
                    chunks: ["taskpane", "vendor"],
                    hash: true
                }),
            new CopyWebpackPlugin({
                patterns: [
                    { from: "./src/public/auth-*.html", to: "../[name][ext]" }, // the dest folder is based on dist/web/scripts
                    { from: "./src/public/assets/*.*", to: "../assets/[name][ext]" },
                    { from: "./src/public/styles/*.*", to: "../styles/[name][ext]" },
                    { from: "./src/settings.*.json", to: "../../[name][ext]" }
                ],
            }),
            new BundleAnalyzerPlugin()

        ],
        // devServer: {
        //     hot: false,
        //     host: "localhost",
        //     port: 9000,
        //     allowedHosts: "all",
        //     client: {
        //         overlay: {
        //             warnings: false,
        //             errors: true
        //         }
        //     },
        //     devMiddleware: {
        //         writeToDisk: true,
        //         stats: {
        //             all: false,
        //             colors: true,
        //             errors: true,
        //             warnings: true,
        //             timings: true,
        //             entrypoints: true
        //         }
        //     }
        // }
    }
];

// if (lint !== false) {
//     config[0].plugins.push(new ESLintPlugin({ extensions: ["ts", "tsx"], failOnError: false, lintDirtyModulesOnly: debug }));
//     config[1].plugins.push(new ESLintPlugin({ extensions: ["ts", "tsx"], failOnError: false, lintDirtyModulesOnly: debug }));
// }

module.exports = config;
