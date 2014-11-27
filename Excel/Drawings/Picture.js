'use strict';

var _ = require('underscore');
var util = require('../util');
var Picture = require('./Picture');

/**
 * @constructor
 */
var Picture = function () {
  this.media = null;
  this.id = _.uniqueId('Picture');
  this.pictureId = util.uniqueId('Picture');
  this.fill = {};
  this.mediaData = null;
};

Picture.prototype = new Drawing();

Picture.prototype.setMedia = function (mediaRef) {
  this.mediaData = mediaRef;
};

Picture.prototype.setDescription = function (description) {
  this.description = description;
};

Picture.prototype.setFillType = function (type) {
  this.fill.type = type;
};

Picture.prototype.setFillConfig = function (config) {
  _.extend(this.fill, config);
};

Picture.prototype.getMediaType = function () {
  return 'image';
};

Picture.prototype.getMediaData = function () {
  return this.mediaData;
};

Picture.prototype.setRelationshipId = function (rId) {
  this.mediaData.rId = rId;
};

Picture.prototype.toXML = function (xmlDoc) {
  var pictureNode = util.createElement(xmlDoc, 'xdr:pic');
            
  var nonVisibleProperties = util.createElement(xmlDoc, 'xdr:nvPicPr');
  
  var nameProperties = util.createElement(xmlDoc, 'xdr:cNvPr', [
      ['id', this.pictureId],
      ['name', this.mediaData.fileName],
      ['descr', this.description || ""]
  ]);
  nonVisibleProperties.appendChild(nameProperties);
  var nvPicProperties = util.createElement(xmlDoc, 'xdr:cNvPicPr');
  nvPicProperties.appendChild(util.createElement(xmlDoc, 'a:picLocks', [
      ['noChangeAspect', '1'],
      ['noChangeArrowheads', '1']
  ]));
  nonVisibleProperties.appendChild(nvPicProperties);
  pictureNode.appendChild(nonVisibleProperties);
  var pictureFill = util.createElement(xmlDoc, 'xdr:blipFill');
  pictureFill.appendChild(util.createElement(xmlDoc, 'a:blip', [
      ['xmlns:r', util.schemas.relationships],
      ['r:embed', this.mediaData.rId]
  ]));
  pictureFill.appendChild(util.createElement(xmlDoc, 'a:srcRect'));
  var stretch = util.createElement(xmlDoc, 'a:stretch');
  stretch.appendChild(util.createElement(xmlDoc, 'a:fillRect'));
  pictureFill.appendChild(stretch);
  pictureNode.appendChild(pictureFill);
  
  var shapeProperties = util.createElement(xmlDoc, 'xdr:spPr', [
      ['bwMode', 'auto']
  ]);
  
  var transform2d = util.createElement(xmlDoc, 'a:xfrm');
  shapeProperties.appendChild(transform2d);
  
  var presetGeometry = util.createElement(xmlDoc, 'a:prstGeom', [
      ['prst', 'rect']
  ]);
  shapeProperties.appendChild(presetGeometry);
  
  pictureNode.appendChild(shapeProperties);

  return this.anchor.toXML(xmlDoc, pictureNode);    
};

module.exports = Picture;