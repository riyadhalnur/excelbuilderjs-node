'use strict';

var express = require('express');
var excel = require('../index.js');
var BasicReport = require('../Template/BasicReport');
//var testdata = require('./testdata.json');
var app = express();

app.get('/', function(req, res) {

  //var data = JSON.parse(testdata);
  //console.log(data);
                
  var basicReport = new BasicReport();
  var columns = [
    {id: 'id', name: 'ID', type: 'number', width: 20},
    {id: 'name', name:'Name', type: 'string', width: 50},
    {id: 'number', name: 'Number', type: 'number', width: 50},
    {id: 'address', name: 'Address', type: 'string', width: 50},
    {id: 'lat', name: 'Lat', type: 'number', width: 15}
  ];
  
  var worksheetData = [
    [
      {value: 'ID', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}, 
      {value: 'Name', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}, 
      {value: 'Number', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}, 
      {value: 'Address', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}, 
      {value: 'Lat', metadata: {style: basicReport.predefinedFormatters.header.id, type: 'string'}}
    ]
  ];
  
  basicReport.setHeader([
    {bold: true, text: 'Generic Report'}, '', ''
  ]);

  basicReport.setData(worksheetData);
  basicReport.setColumns(columns);
  basicReport.setFooter([
    '', '', 'Page &P of &N'
  ]);

  var result = excel.createFile(basicReport.prepare());
  console.log(result);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64');
  res.setHeader('Content-Disposition', 'attachment; filename=' + 'demo.xlsx');
  res.end(result, 'binary');
});

app.listen(3000);
console.log('Listening on port 3000');