'use strict';

var Worksheet = require('./Worksheet');

var requireConfig, worksheet;

// TODO

var console = {
  log: postMessage
};

var start = function (data) {
  worksheet = new Worksheet();
  worksheet.importData(data);
  postMessage({status: 'sharedStrings', data: worksheet.collectSharedStrings()});
};

var onmessage = function (event) {
  var data = event.data;
  if (typeof data == 'object') {
    switch (data.instruction) {
      case "setup":
        requireConfig = data.config;
        importScripts(data.requireJsPath);
        require.config(requireConfig);
        postMessage({status: "ready"});
        break;
      case "start": 
        start(data.data);
        break;
      case "export":
        worksheet.setSharedStringCollection({
            strings: data.sharedStrings
        });
        postMessage({status: "finished", data: worksheet.toXML().toString()});
        break;
    }
  }
};

module.exports = {
  console: console,
  start: start,
  onmessage: onmessage
};