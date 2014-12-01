'use strict';

var express = require('express');
var app = express();
var path = require('path');

var Excel = require('../index.js');
var Drawings = require('../Excel/Drawings');
var Picture = require('../Excel/Drawing/Picture');
var Positioning = require('../Excel/Positioning');
var util = require('../Excel/util');
var BasicReport = require('../Template/BasicReport');

app.get('/', function(req, res) {

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
  ];

  var demoList = demoWorkbook.createWorksheet({name: 'Demo List'});

  demoList.setData(worksheetData);
  demoList.setColumns(columns);

  demoWorkbook.addWorksheet(demoList);

  var result = Excel.createFile(demoWorkbook);
  var data = new Buffer(result, 'base64');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64');
  res.setHeader('Content-Disposition', 'attachment; filename=' + 'demo.xlsx');
  res.end(data);
});

app.get('/table', function(req, res) {

  var data = require('./testdata.json');

  var basicReport = new BasicReport();
  var columns = [
    {id: 'id', name: 'ID', type: 'number', width: 20},
    {id: 'name', name:'Name', type: 'string', width: 50},
    {id: 'price', name: 'Price', type: 'number', style: basicReport.predefinedFormatters.currency.id},
    {id: 'location', name: 'Location', type: 'string'},
    {id: 'startDate', name: 'Start Date', type: 'date', style: basicReport.predefinedFormatters.date.id, width: 15},
    {id: 'endDate', name: 'End Date', type: 'date', style: basicReport.predefinedFormatters.date.id, width: 15}
  ];

  var worksheetData = [
    [
      {value: 'ID', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: 'Name', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: 'Price', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: 'Location', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: 'Start Date', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: 'End Date', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}
    ]
  ].concat(data);

  basicReport.setData(worksheetData);
  basicReport.setColumns(columns);

  var result = Excel.createFile(basicReport.prepare());

  var base64Out = new Buffer(result, 'base64');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64');
  res.setHeader('Content-Disposition', 'attachment; filename=' + 'demo.xlsx');

  res.end(base64Out);
});

app.get('/image', function(req, res) {
  var catWorkbook = Excel.createWorkbook();
  var catList = catWorkbook.createWorksheet({name: 'Cat List'});
  var stylesheet = catWorkbook.getStyleSheet();

  var drawings = new Drawings();
  var catImageData = util.base64Encode(path.resolve(__dirname, 'grumpycat.jpg'));

  var picRef = catWorkbook.addMedia('image', 'grumpycat.jpg', catImageData);

  var catPicture1 = new Picture();
  catPicture1.createAnchor('twoCellAnchor', {
    from: {
      x: 0,
      y: 0
    },
    to: {
      x: 3,
      y: 3
    }
  });

  catPicture1.setMedia(picRef);
  drawings.addDrawing(catPicture1);

  var catPicture2 = new Picture();
  catPicture2.createAnchor('absoluteAnchor', {
    x: Positioning.pixelsToEMUs(300),
    y: Positioning.pixelsToEMUs(300),
    width: Positioning.pixelsToEMUs(300),
    height: Positioning.pixelsToEMUs(300)
  });

  catPicture2.setMedia(picRef);
  drawings.addDrawing(catPicture2);

  var catPicture3 = new Picture();
  catPicture3.createAnchor('oneCellAnchor', {
    x: 1,
    y: 1,
    width: Positioning.pixelsToEMUs(300),
    height: Positioning.pixelsToEMUs(300)
  });

  catPicture3.setMedia(picRef);
  drawings.addDrawing(catPicture3);

  catList.addDrawings(drawings);

  catWorkbook.addDrawings(drawings);
  catWorkbook.addWorksheet(catList);

  console.log(catWorkbook.generateFiles());

  var data = Excel.createFile(catWorkbook);
  var result = new Buffer(data, 'base64');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64');
  res.setHeader('Content-Disposition', 'attachment; filename=' + 'demo.xlsx');
  res.end(result);
});


app.listen(3000);
console.log('Listening on port 3000');