'use strict';

var pixelsToEMUs = function (pixels) {
  return Math.round(pixels * 914400 / 96);
};

module.exports = {
  pixelsToEMUs: pixelsToEMUs
};