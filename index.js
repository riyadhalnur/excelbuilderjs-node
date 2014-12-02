'use strict';

var _ = require('underscore');
var JSZip = require('jszip');

var Drawings = require('./Excel/Drawings');
var Paths = require('./Excel/Paths');
var Positioning = require('./Excel/Positioning');
var RelationshipManager = require('./Excel/RelationshipManager');
var SharedStrings = require('./Excel/SharedStrings');
var StyleSheet = require('./Excel/StyleSheet');
var Table = require('./Excel/Table');
var util = require('./Excel/util');
var Workbook = require('./Excel/Workbook');
var Worksheet = require('./Excel/Worksheet');
//var WorksheetExportWorker = require('./Excel/WorksheetExportWorker');
var XMLDOM = require('./Excel/XMLDOM');
//var Worker = require('./Excel/ZipWorker');

var AbsoluteAnchor = require('./Excel/Drawing/AbsoluteAnchor');
var Chart  = require('./Excel/Drawing/Chart');
var Drawing = require('./Excel/Drawing/Drawing');
var OneCellAnchor = require('./Excel/Drawing/OneCellAnchor');
var Picture = require('./Excel/Drawing/Picture');
var TwoCellAnchor = require('./Excel/Drawing/TwoCellAnchor');

module.exports = {
  Drawings: Drawings,
  Paths: Paths,
  Positioning: Positioning,
  RelationshipManager: RelationshipManager,
  SharedStrings: SharedStrings,
  StyleSheet: StyleSheet,
  Table: Table,
  util: util,
  Worksheet: Worksheet,
  //WorksheetExportWorker: WorksheetExportWorker,
  XMLDOM: XMLDOM,

  AbsoluteAnchor: AbsoluteAnchor,
  Chart: Chart,
  Drawing: Drawing,
  OneCellAnchor: OneCellAnchor,
  Picture: Picture,
  TwoCellAnchor: TwoCellAnchor,
  
  /**
   * Creates a new workbook.
   */
  createWorkbook: function () {
    return new Workbook();
  },
  
  /**
   * Turns a workbook into a downloadable file. 
   * @param {Excel/Workbook} workbook The workbook that is being converted
   * @param {Object} options
   * @param {Boolean} options.base64 Whether to 'return' the generated file as a base64 string
   * @param {Function} options.success The callback function to run after workbook creation is successful.
   * @param {Function} options.error The callback function to run if there is an error creating the workbook.
   * @param {String} options.requireJsPath (Optional) The path to requirejs. Will use the id 'requirejs' to look up the script if not specified.
   */
  createFileAsync: function (workbook, options) {
    // workbook.generateFilesAsync({
    //   success: function (files) {
    //     var worker = new Worker();
    //     worker.addEventListener('message', function (event, data) {
    //       if(event.data.status === 'done') {
    //         options.success(event.data.data);
    //       }
    //     });
    //     worker.postMessage({
    //       files: files,
    //       ziplib: require.toUrl('JSZip'),
    //       base64: (!options || options.base64 !== false)
    //     });
    //   },
    //   error: function () {
    //     options.error();
    //   }
    // });
  },
  
  /**
   * Turns a workbook into a downloadable file.
   * @param {Excel/Workbook} workbook The workbook that is being converted
   * @param {Object} options - options to modify how the zip is created. See http://stuk.github.io/jszip/#doc_generate_options
   */
  createFile: function (workbook, options) {
    var zip = new JSZip();
    var files = workbook.generateFiles();
    
    _.each(files, function (content, path) {
      path = path.substr(1);
      if (path.indexOf('.xml') !== -1 || path.indexOf('.rel') !== -1) {
        zip.file(path, content, {base64: false});
      } else {
        zip.file(path, content, {base64: true, binary: true});
      }
    });

    return zip.generate(_.defaults(options || {}, {
      type: 'base64'
    }));
  }
};