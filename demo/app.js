'use strict';

var express = require('express');
var excel = require('../index.js');
var Table = require('../Excel/Table');
//var testdata = require('./testdata.json');
var app = express();

app.get('/', function(req, res) {

  //var data = JSON.parse(testdata);
  //console.log(data);
  var demoWorkbook = excel.createWorkbook();
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

  var result = excel.createFile(demoWorkbook);
  var data = new Buffer(result, 'base64');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64');
  res.setHeader('Content-Disposition', 'attachment; filename=' + 'demo.xlsx');
  res.end(data);
});

app.get('/table', function(req, res) {
  var artistWorkbook = excel.createWorkbook();
  var stylesheet = artistWorkbook.getStyleSheet();
  var albumList = artistWorkbook.createWorksheet({name: 'Album List'});

   var themeColor = stylesheet.createFormat({
    font: {
      bold: true,
      color: red
    }
  });

  var originalData = [
    [
      {value: 'ID', metadata: {style: themeColor.id, type: 'string'}}, 
      {value: 'Name', metadata: {style: themeColor.id, type: 'string'}}, 
      {value: 'Number', metadata: {style: themeColor.id, type: 'string'}}, 
      {value: 'Address', metadata: {style: themeColor.id, type: 'string'}}, 
      {value: 'Lat', metadata: {style: themeColor.id, type: 'string'}}
    ]
  ];

  var albumTable = new Table();
  albumTable.styleInfo.themeStyle = "TableStyleDark2";

  artistWorkbook.addWorksheet(albumList);

  albumList.addTable(albumTable);
  artistWorkbook.addTable(albumTable);
  

  //Table columns are required, even if headerRowCount is zero. The name of the column also must match the
  //data in the column cell that is the header - keep this in mind for localization
  albumTable.setTableColumns([
    'Artist',
    'Album',
    'Price'
  ]);

  albumList.setData(originalData);
  albumTable.setReferenceRange([1, 1], [3, originalData.length]); //X/Y position where the table starts and stops.

  var result = excel.createFile(artistWorkbook);
  var data = new Buffer(result, 'base64');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64');
  res.setHeader('Content-Disposition', 'attachment; filename=' + 'demo.xlsx');
  res.end(data);
});

app.listen(3000);
console.log('Listening on port 3000');