'use strict';

var util = require('../util');

/**
 * 
 * @param {Object} config
 * @param {Number} config.x X offset in EMU's
 * @param {Number} config.y Y offset in EMU's
 * @param {Number} config.width Width in EMU's
 * @param {Number} config.height Height in EMU's
 * @constructor
 */
var AbsoluteAnchor = function (config) {
  this.x = null;
  this.y = null;
  this.width = null;
  this.height = null;

  if (config) {
    this.setPos(config.x, config.y);
    this.setDimensions(config.width, config.height);
  }
};

/**
 * Sets the X and Y offsets.
 *
 * @param {Number} x
 * @param {Number} y
 * @returns {undefined}
 */
AbsoluteAnchor.prototype.setPos = function (x, y) {
  this.x = x;
  this.y = y;
};

/**
 * Sets the width and height of the image.
 *
 * @param {Number} width
 * @param {Number} height
 * @returns {undefined}
 */
AbsoluteAnchor.prototype.setDimensions = function (width, height) {
  this.width = width;
  this.height = height;
};

AbsoluteAnchor.prototype.toXML = function (xmlDoc, content) {
  var root = util.createElement(xmlDoc, 'xdr:absoluteAnchor');
  var pos = util.createElement(xmlDoc, 'xdr:pos');
  pos.setAttribute('x', this.x);
  pos.setAttribute('y', this.y);
  root.appendChild(pos);

  var dimensions = util.createElement(xmlDoc, 'xdr:ext');
  dimensions.setAttribute('cx', this.width);
  dimensions.setAttribute('cy', this.height);
  root.appendChild(dimensions);

  root.appendChild(content);

  root.appendChild(util.createElement(xmlDoc, 'xdr:clientData'));

  return root;
};

module.exports = AbsoluteAnchor;