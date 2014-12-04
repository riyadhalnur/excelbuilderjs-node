'use strict';

var should = require('should');
var Excel = require('../index');
var fs = require('fs');
var testData = require('../demo/testdata.json');

var Drawings = require('../Excel/Drawings');
var Paths = require('../Excel/Paths');
var Positioning = require('../Excel/Positioning');
var RelationshipManager = require('../Excel/RelationshipManager');
var SharedStrings = require('../Excel/SharedStrings');
var StyleSheet = require('../Excel/StyleSheet');
var Table = require('../Excel/Table');
var util = require('../Excel/util');
var Workbook = require('../Excel/Workbook');
var Worksheet = require('../Excel/Worksheet');
//var WorksheetExportWorker = require('../Excel/WorksheetExportWorker');
var XMLDOM = require('../Excel/XMLDOM');
//var Worker = require('../Excel/ZipWorker');

var AbsoluteAnchor = require('../Excel/Drawing/AbsoluteAnchor');
var Chart  = require('../Excel/Drawing/Chart');
var Drawing = require('../Excel/Drawing/Drawing');
var OneCellAnchor = require('../Excel/Drawing/OneCellAnchor');
var Picture = require('../Excel/Drawing/Picture');
var TwoCellAnchor = require('../Excel/Drawing/TwoCellAnchor');

describe('Excel builder library', function () {

  describe('Create Workbook', function () {
    it('should create a new workbook', function () {
      var testWorkbook = Excel.createWorkbook();
      
      testWorkbook.should.be.instanceOf(Workbook);
    });
  });

  describe('Create File', function () {
    it('should create a base64 encoded file for a given workbook', function () {
      var testWorkbook = Excel.createWorkbook();
      var options = {};

      var result = Excel.createFile(testWorkbook, options);
      result.should.match(/^([0-9a-zA-Z+/]{4})*(([0-9a-zA-Z+/]{2}==)|([0-9a-zA-Z+/]{3}=))?$/);
    });

    it('should create and xlsx file', function () {
      var demoWorkbook = Excel.createWorkbook();
      var stylesheet = demoWorkbook.getStyleSheet();

      var themeColor = stylesheet.createFormat({
        font: {
          bold: true,
          color: {theme: 3}
        }
      });

      var columns = [
        {id: 'id', name: 'ID', type: 'number', width: 10},
        {id: 'name', name:'Name', type: 'string', width: 50},
        {id: 'number', name: 'Number', type: 'number', width: 50},
        {id: 'address', name: 'Address', type: 'string', width: 50},
        {id: 'lat', name: 'Lat', type: 'number', width: 15}
      ];

      var worksheetData = [
        [
          {value: 'ID', metadata: {style: themeColor.id, type: 'string'}}, 
          {value: 'Name', metadata: {style: themeColor.id, type: 'string'}}, 
          {value: 'Number', metadata: {style: themeColor.id, type: 'string'}}, 
          {value: 'Address', metadata: {style: themeColor.id, type: 'string'}}, 
          {value: 'Lat', metadata: {style: themeColor.id, type: 'string'}}
        ]
      ].concat(testData);

      var demoList = demoWorkbook.createWorksheet({name: 'Demo List'});

      demoList.setData(worksheetData);
      demoList.setColumns(columns);

      demoWorkbook.addWorksheet(demoList);

      var result = Excel.createFile(demoWorkbook);
      var data = new Buffer(result, 'base64');

      fs.writeFileSync('test.xlsx', data);
    });
  });

});