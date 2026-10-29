// this module allows the specs to load modules from the source tree

const libs = {};
const basePath = '../../lib/';
libs.exceljs = require('../../lib/exceljs.nodejs');

module.exports = function verquire(path) {
  if (!libs[path]) {
    libs[path] = require(basePath + path);
  }
  return libs[path];
};
