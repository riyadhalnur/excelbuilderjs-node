'use strict';

var _ = require('underscore');
var RelationshipManager = require('./RelationshipManager');
var util = require('./util');

var Worksheet = function (config) {
  this.relations = null;
  this.columnFormats = [];
  this.data = [];
  this.mergedCells = [];
  this.columns = [];
  this._headers = [];
  this._footers = [];
  this._tables = [];
  this._drawings = [];
  this.initialize(config);
};

Worksheet.prototype.initialize = function (config) {
  config = config || {};
  this.name = config.name;
  this.id = _.uniqueId('Worksheet');
  this._timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;

  if (config.columns) {
    this.setColumns(config.columns);
  }
  
  this.relations = new RelationshipManager();
};

Worksheet.prototype.exportData = function () {
  return {
    relations: this.relations.exportData(),
    columnFormats: this.columnFormats,
    data: this.data,
    columns: this.columns,
    mergedCells: this.mergedCells,
    _headers: this._headers,
    _footers: this._footers,
    _tables: this._tables,
    name: this.name,
    id: this.id
  };
};

Worksheet.prototype.importData = function (data) {
  this.relations.importData(data.relations);
  delete data.relations;
  _.extend(this, data);
};

Worksheet.prototype.setSharedStringCollection = function (stringCollection) {
  this.sharedStrings = stringCollection;
};

Worksheet.prototype.addTable = function (table) {
  this._tables.push(table);
  this.relations.addRelation(table, 'table');
};

Worksheet.prototype.addDrawings = function (table) {
  this._drawings.push(table);
  this.relations.addRelation(table, 'drawingRelationship');
};

Worksheet.prototype.setHeader = function (headers) {
  if(!_.isArray(headers)) {
    throw "Invalid argument type - setHeader expects an array of three instructions";
  }
  this._headers = headers;
};

Worksheet.prototype.setFooter = function (footers) {
  if(!_.isArray(footers)) {
    throw "Invalid argument type - setFooter expects an array of three instructions";
  }
  this._footers = footers;
};

Worksheet.prototype.compilePageDetailPackage = function (data) {
  data = data || "";
  return [
  "&L", this.compilePageDetailPiece(data[0] || ""),
  "&C", this.compilePageDetailPiece(data[1] || ""),
  "&R", this.compilePageDetailPiece(data[2] || "")
  ].join('');
};

Worksheet.prototype.compilePageDetailPiece = function (data) {
  if (_.isString(data)) {
    return '&"-,Regular"'.concat(data);
  }

  if(_.isObject(data) && !_.isArray(data)) { 
    var string = "";

    if (data.font || data.bold) {
      var weighting = data.bold ? "Bold" : "Regular";
      string += '&"' + (data.font || '-');
      string += ',' + weighting + '"';
    } else {
      string += '&"-,Regular"';
    }

    if (data.underline) {
      string += "&U";
    }

    if (data.fontSize) {
      string += "&"+data.fontSize;
    }

    string += data.text;
    
    return string;
  }
  
  if (_.isArray(data)) {
    var self = this;
    return _.reduce(data, function (m, v) {
      return m.concat(self.compilePageDetailPiece(v));
    }, "");
  }
};

Worksheet.prototype.exportHeader = function (doc) {
  var oddHeader = doc.createElement('oddHeader');
  oddHeader.appendChild(doc.createTextNode(this.compilePageDetailPackage(this._headers)));
  return oddHeader;
};

Worksheet.prototype.exportFooter = function (doc) {
  var oddFooter = doc.createElement('oddFooter');
  oddFooter.appendChild(doc.createTextNode(this.compilePageDetailPackage(this._footers)));
  return oddFooter;
};

Worksheet.prototype._buildCache = function (doc) {
  var numberNode = doc.createElement('c');
  var value = doc.createElement('v');
  value.appendChild(doc.createTextNode("--temp--"));
  numberNode.appendChild(value);
  
  var formulaNode = doc.createElement('c');
  var formulaValue = doc.createElement('f');
  formulaValue.appendChild(doc.createTextNode("--temp--"));
  formulaNode.appendChild(formulaValue);
  
  var stringNode = doc.createElement('c');
  stringNode.setAttribute('t', 's');
  var stringValue = doc.createElement('v');
  stringValue.appendChild(doc.createTextNode("--temp--"));
  stringNode.appendChild(stringValue);
  
  
  return {
    number: numberNode,
    date: numberNode,
    string: stringNode,
    formula: formulaNode
  };
};

Worksheet.prototype.collectSharedStrings = function () {
  var data = this.data;
  var maxX = 0;
  var strings = {};

  for (var row = 0, l = data.length; row < l; row++) {
    var dataRow = data[row];
    var cellCount = dataRow.length;
    maxX = cellCount > maxX ? cellCount : maxX;

    for (var c = 0; c < cellCount; c++) {
      var cellValue = dataRow[c];
      var metadata = cellValue && cellValue.metadata || {};

      if (cellValue && typeof cellValue === 'object') {
        cellValue = cellValue.value;
      }
      
      if (!metadata.type) {
        if (typeof cellValue === 'number') {
            metadata.type = 'number';
        }
      }

      if (metadata.type === "text" || !metadata.type) {
        if (typeof strings[cellValue] === 'undefined') {
          strings[cellValue] = true;
        }
      }
    }
  }

  return _.keys(strings);
};

Worksheet.prototype.toXML = function () {
  var data = this.data;
  var columns = this.columns || [];
  var doc = util.createXmlDoc(util.schemas.spreadsheetml, 'worksheet');
  var worksheet = doc.documentElement;
  var i, l, row;
  worksheet.setAttribute('xmlns:r', util.schemas.relationships);
  worksheet.setAttribute('xmlns:mc', util.schemas.markupCompat);
  
  var maxX = 0;
  var sheetData = util.createElement(doc, 'sheetData');
  
  var cellCache = this._buildCache(doc);
  
  for (row = 0, l = data.length; row < l; row++) {
    var dataRow = data[row];
    var cellCount = dataRow.length;
    maxX = cellCount > maxX ? cellCount : maxX;
    var rowNode = doc.createElement('row');
    
    for (var c = 0; c < cellCount; c++) {
      columns[c] = columns[c] || {};
      var cellValue = dataRow[c];
      var cell, metadata = cellValue && cellValue.metadata || {};

      if (cellValue && typeof cellValue === 'object') {
        cellValue = cellValue.value;
      }

      if (!metadata.type) {
        if (typeof cellValue === 'number') {
          metadata.type = 'number';
        }
      }

      switch(metadata.type) {
        case "number":
          cell = cellCache.number.cloneNode(true);
          cell.firstChild.firstChild.nodeValue = cellValue;
          break;
        case "date":
          cell = cellCache.date.cloneNode(true);
          cell.firstChild.firstChild.nodeValue = 25569.0 + ((cellValue - this._timezoneOffset)  / (60 * 60 * 24 * 1000));
          break;
        case "formula":
          cell = cellCache.formula.cloneNode(true);
          cell.firstChild.firstChild.nodeValue = cellValue;
          break;
        case "text":
          /*falls through*/
        default:
          var id;
          if (typeof this.sharedStrings.strings[cellValue] !== 'undefined') {
            id = this.sharedStrings.strings[cellValue];
          } else {
            id = this.sharedStrings.addString(cellValue);
          }
          cell = cellCache.string.cloneNode(true);
          cell.firstChild.firstChild.nodeValue = id;
          break;
      }

      if (metadata.style) {
        cell.setAttribute('s', metadata.style);
      }

      cell.setAttribute('r', util.positionToLetterRef(c + 1, row + 1));
      rowNode.appendChild(cell);
    }

    rowNode.setAttribute('r', row + 1);
    sheetData.appendChild(rowNode);
  } 
  
  if (maxX !== 0) {
    worksheet.appendChild(util.createElement(doc, 'dimension', [
        ['ref',  util.positionToLetterRef(1, 1) + ':' + util.positionToLetterRef(maxX, data.length)]
    ]));
  } else {
    worksheet.appendChild(util.createElement(doc, 'dimension', [
        ['ref',  util.positionToLetterRef(1, 1)]
    ]));
  }
  
  if (this.columns.length) {
    worksheet.appendChild(this.exportColumns(doc));
  }
  worksheet.appendChild(sheetData);

  // 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
  // with Microsoft Excel (2007, 2013)
  if (this.mergedCells.length > 0) {
    var mergeCells = doc.createElement('mergeCells');
    for (i = 0, l = this.mergedCells.length; i < l; i++) {
      var mergeCell = doc.createElement('mergeCell');
      mergeCell.setAttribute('ref', this.mergedCells[i][0] + ':' + this.mergedCells[i][1]);
      mergeCells.appendChild(mergeCell);
    }
    worksheet.appendChild(mergeCells);
  }
  
  this.exportPageSettings(doc, worksheet);
  
  if (this._headers.length > 0 || this._footers.length > 0) {
    var headerFooter = doc.createElement('headerFooter');
    if (this._headers.length > 0) {
      headerFooter.appendChild(this.exportHeader(doc));
    }
    if (this._footers.length > 0) {
      headerFooter.appendChild(this.exportFooter(doc));
    }
    worksheet.appendChild(headerFooter);
  }
  
  if (this._tables.length > 0) {
    var tables = doc.createElement('tableParts');
    tables.setAttribute('count', this._tables.length);
    for (i = 0, l = this._tables.length; i < l; i++) {
      var table = doc.createElement('tablePart');
      table.setAttribute('r:id', this.relations.getRelationshipId(this._tables[i]));
      tables.appendChild(table);
    }
    worksheet.appendChild(tables);
  }

  // the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
  // to issue with Microsoft Excel (2007, 2013)
  for (i = 0, l = this._drawings.length; i < l; i++) {
    var drawing = doc.createElement('drawing');
    drawing.setAttribute('r:id', this.relations.getRelationshipId(this._drawings[i]));
    worksheet.appendChild(drawing);
  }

  return doc;
};

Worksheet.prototype.exportColumns = function (doc) {
  var cols = util.createElement(doc, 'cols');
  for (var i = 0, l = this.columns.length; i < l; i++) {
    var cd = this.columns[i];
    var col = util.createElement(doc, 'col', [
        ['min', cd.min || i + 1],
        ['max', cd.max || i + 1]
    ]);

    if (cd.hidden) {
      col.setAttribute('hidden', 1);
    }

    if (cd.bestFit) {
      col.setAttribute('bestFit', 1);
    }

    if (cd.customWidth || cd.width) {
      col.setAttribute('customWidth', 1);
    }

    if (cd.width) {
      col.setAttribute('width', cd.width);
    } else {
      col.setAttribute('width', 9.140625);
    }
    
    cols.appendChild(col);
  }

  return cols;
};

Worksheet.prototype.exportPageSettings = function () {
  if (this._orientation) {
    worksheet.appendChild(util.createElement(doc, 'pageSetup', [
      ['orientation', this._orientation]
    ]));
  }
};

Worksheet.prototype.setPageOrientation = function (orientation) {
  this._orientation = orientation;
};

Worksheet.prototype.setColumns = function (columns) {
  this.columns = columns;
};

Worksheet.prototype.setData = function (data) {
  this.data = data;
};

Worksheet.prototype.mergeCells = function (cell1, cell2) {
  this.mergedCells.push([cell1, cell2]);
};

Worksheet.prototype.setColumnFormats = function (columnFormats) {
  this.columnFormats = columnFormats;
};

module.exports = {
  Worksheet: Worksheet
};