'use strict';

var pixelToEMUs = function (pixels) {
  return Math.round(pixels * 914400 / 96);
};

module.exports = pixelToEMUs;