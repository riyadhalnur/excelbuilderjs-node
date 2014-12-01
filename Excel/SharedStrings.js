'use strict';

var _ = require('underscore');
var util = require('./util');

var sharedStrings = function () {
  this.strings = {};
  this.stringArray = [];
  this.id = _.uniqueId('SharedStrings');
};

/**
 * Adds a string to the shared string file, and returns the ID of the
 * string which can be used to reference it in worksheets.
 *
 * @param string {String}
 * @return int
 */
sharedStrings.prototype.addString = function (string) {
  this.strings[string] = this.stringArray.length;
  this.stringArray[this.stringArray.length] = string;

  return this.strings[string];
};

sharedStrings.prototype.exportData = function () {
  return this.strings;
};

sharedStrings.prototype.toXML = function () {
  var doc = util.createXmlDoc(util.schemas.spreadsheetml, 'sst');
  var sharedStringTable = doc.documentElement;
  this.stringArray.reverse();
  var l = this.stringArray.length;
  sharedStringTable.setAttribute('count', l);
  sharedStringTable.setAttribute('uniqueCount', l);

  var template = doc.createElement('si');
  var templateValue = doc.createElement('t');
  templateValue.appendChild(doc.createTextNode('--placeholder--'));
  template.appendChild(templateValue);
  var strings = this.stringArray;

  while (l--) {
    var clone = template.cloneNode(true);
    clone.firstChild.firstChild.nodeValue = strings[l];
    sharedStringTable.appendChild(clone);
  }

  return doc;
};

module.exports = sharedStrings;