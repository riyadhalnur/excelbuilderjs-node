'use strict';

var express = require('express');
var excel = require('../index.js');
var app = express();

app.get('/', function(req, res){
  res.setHeader('Content-Type', 'application/vnd.openxmlformats');
  res.setHeader("Content-Disposition", "attachment; filename=" + "demo.xlsx");
  res.end(result, 'binary');
});

app.listen(3000);
console.log('Listening on port 3000');