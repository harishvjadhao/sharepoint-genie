"use strict";

const build = require("@microsoft/sp-build-web");
const webpack = require("webpack");
const fs = require("fs");

//possible values qa, uat, prod
const env = "prod";

const prodConfigPath = `${process
  .cwd()
  .replace(/\\/g, "/")}/${env}.config.json`;

const prodConfigExists = fs.existsSync(prodConfigPath);

const APP_CONFIG_VALUES = prodConfigExists
  ? require(prodConfigPath)
  : undefined;

console.log("APP_CONFIG");
console.log(JSON.stringify(APP_CONFIG_VALUES, null, 4));

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);
build.addSuppression(/Warning/gi);

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    generatedConfiguration.plugins.push(
      new webpack.DefinePlugin({
        APP_CONFIG: JSON.stringify(APP_CONFIG_VALUES),
      })
    );

    generatedConfiguration.externals = generatedConfiguration.externals.filter(
      (e) => !["react", "react-dom"].includes(e)
    );
    return generatedConfiguration;
  },
});

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set("serve", result.get("serve-deprecated"));

  return result;
};

build.initialize(require("gulp"));
