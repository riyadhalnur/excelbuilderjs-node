'use strict';

var _ = require('underscore');
var RelationshipManager = require('./RelationshipManager');
var util = require('./util');

var Drawings = function () {
  this.drawings = [];
  this.relations = new RelationshipManager();
  this.id = _.uniqueId('Drawings');
};

/**
 * Adds a drawing (more likely a subclass of a Drawing) to the 'Drawings' for a particular worksheet.
 *
 * @param {Drawing} drawing
 * @returns {undefined}
 */
Drawings.prototype.addDrawing = function (drawing) {
  this.drawings.push(drawing);
};

Drawings.prototype.getCount = function () {
  return this.drawings.length;
};

Drawings.prototype.toXML = function () {
  var doc = util.createXmlDoc(util.schemas.spreadsheetDrawing, 'xdr:wsDr');
  var drawings = doc.documentElement;
  drawings.setAttribute('xmlns:a', util.schemas.drawing);
  drawings.setAttribute('xmlns:xdr', util.schemas.spreadsheetDrawing);

  for(var i = 0, l = this.drawings.length; i < l; i++) {

    var rId = this.relations.getRelationshipId(this.drawings[i].getMediaData());
    if(!rId) {
      rId = this.relations.addRelation(this.drawings[i].getMediaData(), this.drawings[i].getMediaType()); //chart
    }
    this.drawings[i].setRelationshipId(rId);
    drawings.appendChild(this.drawings[i].toXML(doc));
  }

  return doc;
};

module.exports = Drawings;