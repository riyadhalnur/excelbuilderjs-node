'use strict';

var express = require('express');
var excel = require('../index.js');
var Table = require('../Excel/Table');
var BasicReport = require('../Template/BasicReport');

var app = express();

app.get('/', function(req, res) {

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

  var data = require('./testdata.json');

  var basicReport = new BasicReport();
  var columns = [
    {id: 'id', name: "ID", type: 'number', width: 20},
    {id: 'name', name:"Name", type: 'string', width: 50},
    {id: 'price', name: "Price", type: 'number', style: basicReport.predefinedFormatters.currency.id},
    {id: 'location', name: "Location", type: 'string'},
    {id: 'startDate', name: "Start Date", type: 'date', style: basicReport.predefinedFormatters.date.id, width: 15},
    {id: 'endDate', name: "End Date", type: 'date', style: basicReport.predefinedFormatters.date.id, width: 15}
  ];

  var worksheetData = [
    [
      {value: "ID", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: "Name", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: "Price", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: "Location", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: "Start Date", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}},
      {value: "End Date", metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}
    ]
  ].concat(data);

  basicReport.setData(worksheetData);
  basicReport.setColumns(columns);

  var result = excel.createFile(basicReport.prepare());

  var base64Out = new Buffer(result, 'base64');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64');
  res.setHeader('Content-Disposition', 'attachment; filename=' + 'demo.xlsx');

  res.end(base64Out);
});

app.listen(3000);
console.log('Listening on port 3000');