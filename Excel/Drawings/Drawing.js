'use strict';

var AbsoluteAnchor = require('./AbsoluteAnchor');
var OneCellAnchor = require('./OneCellAnchor');
var TwoCellAnchor = require('./TwoCellAnchor');

/**
 * @constructor
 */
var Drawing = function () {
  this.id = _.uniqueId('Drawing');
};

/**
 * 
 * @param {String} type Can be 'absoluteAnchor', 'oneCellAnchor', or 'twoCellAnchor'. 
 * @param {Object} config Shorthand - pass the created anchor coords that can normally be used to construct it.
 * @returns {Anchor}
 */
Drawing.prototype.createAnchor = function (type, config) {
  config = config || {};
  config.drawing = this;

  switch(type) {
    case 'absoluteAnchor': 
      this.anchor = new AbsoluteAnchor(config);
      break;
    case 'oneCellAnchor':
      this.anchor = new OneCellAnchor(config);
      break;
    case 'twoCellAnchor':
      this.anchor = new TwoCellAnchor(config);
      break;
  }

  return this.anchor;
};

module.exports = {
  Drawing: Drawing
};