'use strict';

const build = require('@microsoft/sp-build-web');
const webpack = require("webpack");
const getClientEnvironment = require("./process-env");

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.configureWebpack.mergeConfig({
  additionalConfiguration: cfg => {
    let pluginDefine = null;
    for (var i = 0; i < cfg.plugins.length; i++) {
      var plugin = cfg.plugins[i];
      if (plugin instanceof webpack.DefinePlugin) {
        pluginDefine = plugin;
      }
    }
 
    const currentEnv = getClientEnvironment().stringified;
    console.log('currentEnv', currentEnv);
 
    if (pluginDefine) {
      pluginDefine.definitions = { ...pluginDefine.definitions, ...currentEnv };
      console.log('pluginDefine', pluginDefine);
    } else {
      cfg.plugins.push(new webpack.DefinePlugin(currentEnv));
    }
    return cfg;
  }
});

build.initialize(require('gulp'));
