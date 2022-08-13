'use strict';

const fs = require('fs');
const path = require('path');

const NODE_ENV = process.NODE_ENV || 'dev';
const dotEnvPath = path.resolve(process.cwd(), '.env');

const dotenvFiles = [
  `${dotEnvPath}.${NODE_ENV}.local`,
  `${dotEnvPath}.${NODE_ENV}`,
  NODE_ENV !== 'test' && `${dotEnvPath}.local`,
  dotEnvPath
].filter(Boolean);

let parsedEnv = {};

dotenvFiles.forEach(dotenvFile => {
  if (fs.existsSync(dotenvFile)) {
    const { expand } = require("dotenv-expand")
    const { parsed } = expand(require('dotenv').config({
      path: dotenvFile
    }));
    parsedEnv = {...parsedEnv, ...parsed};
  }
});

function getClientEnvironment() {
  const stringified = {
    'process.env': Object
      .keys(parsedEnv)
      .reduce((env, key) => {
        env[key] = JSON.stringify(parsedEnv[key]);
        return env;
      }, {})
  };
  return { stringified };
}

module.exports = getClientEnvironment;