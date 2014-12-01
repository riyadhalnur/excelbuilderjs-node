'use strict';

var should = require('should');
var Excel = require('../index');
var Workbook = require('../Excel/Workbook');

describe('Excel builder library', function () {

  describe('.createWorkbook', function () {
    it('should create a new workbook', function () {
      var testWorkbook = Excel.createWorkbook();
      
      testWorkbook.should.be.instanceOf(Workbook);
    });
  });

});