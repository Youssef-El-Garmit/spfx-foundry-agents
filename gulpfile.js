'use strict';

const path = require('path');
const build = require('@microsoft/sp-build-web');
const webpack = require('webpack');

const emptyNodeModule = path.resolve(__dirname, 'src/webparts/foundryAgents/emptyNodeModule.js');

// Allow @azure/ai-projects 2.0 to bundle in browser (@azure/identity uses node:fs / node:path)
build.configureWebpack.mergeConfig({
  additionalConfiguration: (config) => {
    config.resolve = config.resolve || {};
    config.resolve.fallback = Object.assign({}, config.resolve.fallback || {}, {
      fs: false,
      path: false
    });
    config.plugins = config.plugins || [];
    config.plugins.push(
      new webpack.NormalModuleReplacementPlugin(/^node:fs$/, emptyNodeModule),
      new webpack.NormalModuleReplacementPlugin(/^node:fs\/promises$/, emptyNodeModule),
      new webpack.NormalModuleReplacementPlugin(/^node:path$/, emptyNodeModule)
    );
    return config;
  }
});

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

build.initialize(require('gulp'));
