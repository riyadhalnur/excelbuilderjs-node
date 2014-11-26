'use strict';

var _ = require('underscore');
var util = require('./util');

/**
 * 
 * @param {Object} config
 * @param {Number} config.from The x and y offset from where the picture will start
 * @param {Number} config.to The x and y offset to where the picture will end
 * @constructor
 */
var TwoCellAnchor = function (config) {
  this.from = {xOff: 0, yOff: 0};
  this.to = {xOff: 0, yOff: 0};

  if(config) {
    this.setFrom(config.from.x, config.from.y, config.to.xOff, config.to.yOff);
    this.setTo(config.to.x, config.to.y, config.to.xOff, config.to.yOff);
  }
};

TwoCellAnchor.prototype.setFrom = function (x, y, xOff, yOff) {
  this.from.x = x; 
  this.from.y = y;

  if(xOff !== undefined) {
    this.from.xOff = xOff;
  }

  if(yOff !== undefined) {
    this.from.yOff = xOff;
  }
};

TwoCellAnchor.prototype.setTo = function (x, y, xOff, yOff) {
  this.to.x = x; 
  this.to.y = y;

  if(xOff !== undefined) {
    this.to.xOff = xOff;
  }

  if(yOff !== undefined) {
    this.to.yOff = xOff;
  }
};

TwoCellAnchor.prototype.toXML = function (xmlDoc, content) {
  var root = util.createElement(xmlDoc, 'xdr:twoCellAnchor');
            
  var from = util.createElement(xmlDoc, 'xdr:from');
  var fromCol = util.createElement(xmlDoc, 'xdr:col');
  fromCol.appendChild(xmlDoc.createTextNode(this.from.x));
  var fromColOff = util.createElement(xmlDoc, 'xdr:colOff');
  fromColOff.appendChild(xmlDoc.createTextNode(this.from.xOff));
  var fromRow = util.createElement(xmlDoc, 'xdr:row');
  fromRow.appendChild(xmlDoc.createTextNode(this.from.y));
  var fromRowOff = util.createElement(xmlDoc, 'xdr:rowOff');
  fromRowOff.appendChild(xmlDoc.createTextNode(this.from.yOff));
  
  from.appendChild(fromCol);
  from.appendChild(fromColOff);
  from.appendChild(fromRow);
  from.appendChild(fromRowOff);
  
  var to = util.createElement(xmlDoc, 'xdr:to');
  var toCol = util.createElement(xmlDoc, 'xdr:col');
  toCol.appendChild(xmlDoc.createTextNode(this.to.x));
  var toColOff = util.createElement(xmlDoc, 'xdr:colOff');
  toColOff.appendChild(xmlDoc.createTextNode(this.from.xOff));
  var toRow = util.createElement(xmlDoc, 'xdr:row');
  toRow.appendChild(xmlDoc.createTextNode(this.to.y));
  var toRowOff = util.createElement(xmlDoc, 'xdr:rowOff');
  toRowOff.appendChild(xmlDoc.createTextNode(this.from.yOff));
  
  to.appendChild(toCol);
  to.appendChild(toColOff);
  to.appendChild(toRow);
  to.appendChild(toRowOff);
  
  
  root.appendChild(from);
  root.appendChild(to);
  
  root.appendChild(content);
  
  root.appendChild(util.createElement(xmlDoc, 'xdr:clientData'));
  return root;
};

module.exports = {
  TwoCellAnchor: TwoCellAnchor
};