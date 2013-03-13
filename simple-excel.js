// SimpleExcel.js v0.0.1
// Client-side parser & writer for Excel file formats (CSV, XML, XLSX)
// https://github.com/faisalman/simple-excel-js
// 
// Copyright © 2013 Faisalman <fyzlman@gmail.com>
// Dual licensed under GPLv2 & MIT

(function (window, undefined) {

    'use strict';

    ///////////////////////
    // Constants & Helpers
    ///////////////////////

    var Char = {
        COMMA           : ',',
        RETURN          : '\r',
        NEWLINE         : '\n',
        SEMICOLON       : ';',
        TAB             : '\t'
    };
    
    var DataType = {
        CURRENCY    : 'CURRENCY',
        DATETIME    : 'DATETIME',
        FORMULA     : 'FORMULA',
        LOGICAL     : 'LOGICAL',
        NUMBER      : 'NUMBER',
        TEXT        : 'TEXT'
    };

    var Exception = {    
        CELL_NOT_FOUND              : 'CELL_NOT_FOUND',
        COLUMN_NOT_FOUND            : 'COLUMN_NOT_FOUND',
        ERROR_READING_FILE          : 'ERROR_READING_FILE',
        ERROR_WRITING_FILE          : 'ERROR_WRITING_FILE',
        FILE_NOT_FOUND              : 'FILE_NOT_FOUND',
        FILE_EXTENSION_MISMATCH     : 'FILE_EXTENSION_MISMATCH',
        FILETYPE_NOT_SUPPORTED      : 'FILETYPE_NOT_SUPPORTED',
        INVALID_DOCUMENT_FORMAT     : 'INVALID_DOCUMENT_FORMAT',
        INVALID_DOCUMENT_NAMESPACE  : 'INVALID_DOCUMENT_NAMESPACE',
        MALFORMED_JSON              : 'MALFORMED_JSON',
        ROW_NOT_FOUND               : 'ROW_NOT_FOUND',
        UNIMPLEMENTED_METHOD        : 'UNIMPLEMENTED_METHOD',
        UNKNOWN_ERROR               : 'UNKNOWN_ERROR',
        UNSUPPORTED_BROWSER         : 'UNSUPPORTED_BROWSER'
    };
    
    var Format = {        
        CSV     : 'CSV',
        TSV     : 'TSV',
        XLSX    : 'XLSX',
        XML     : 'XML'
    };
    
    var MIMEType = {
        CSV     : 'text/csv',
        TSV     : 'text/tsv'
    };

    var Regex = {
        FILENAME    : /.*\./g,
        LINEBREAK   : /\r\n/g
    };

    var Utils = {
        getFiletype : function (filename) {
            return filename.replace(Regex.FILENAME, '');
        },
        isEqual     : function (str1, str2, ignoreCase) {
            return ignoreCase ? str1.toLowerCase() == str2.toLowerCase() : str1 == str2;
        },
        isSupportedBrowser : function () {
            return !![].forEach 
                && !!window.FileReader;
        }
    };
    
    /////////////////////////////
    // Spreadsheet Constructors
    ////////////////////////////

    var Cell = function (value, datatype) {
        this.value = value;
        this.dataType = datatype || DataType.TEXT;
        this.toString = function () {
            return value.toString();
        }
    };
        
    var Records = function () {};
    Records.prototype = new Array();
    Records.prototype.getCell = function (colNum, rowNum) {
        return this[rowNum - 1][colNum - 1];
    };
    Records.prototype.getColumn = function (colNum) {        
        var col = [];
        this.forEach(function (el, i) {
            col.push(el[colNum - 1]);
        });
        return col;
    };
    Records.prototype.getRow = function (rowNum) {
        return this[rowNum - 1];
    };
    
    var Sheet = function (number) {
        this.number = number;
        this.cells = new Records();
    };
    Sheet.prototype.addRow = function (array) {
        this.cells.push(array);
    };
    
    /////////////
    // Parsers
    ////////////

    var BaseParser = function () {};
    BaseParser.prototype = {
        _filetype   : undefined,
        _sheet      : [],
        getSheet    : function (number) {
            var number = number || 1;
            return this._sheet[number - 1].cells;
        },
        loadFile    : function (fileEl, callback) {
            var self = this;
            fileEl.addEventListener('change', function (e) {
                var file = e.target.files[0];
                var filetype = Utils.getFiletype(file.name);
                if (Utils.isEqual(filetype, self._filetype, true)) {
                    var reader = new FileReader();
                    reader.onload = function () {
                        self.loadString(this.result, 0);
                        callback.apply(self, e);
                    };
                    reader.readAsText(file);
                } else {
                    throw Exception.FILE_EXTENSION_MISMATCH;
                }
            }, false);
            return self;
        },
        loadString  : function (string, sheetnum) {
            throw Exception.UNIMPLEMENTED_METHOD;
        }
    };

    var CSVParser = function () {};
    CSVParser.prototype = new BaseParser();
    CSVParser.prototype._delimiter = Char.COMMA;
    CSVParser.prototype._filetype = Format.CSV;
    CSVParser.prototype.loadString = function (string, sheetnum) {
        // TODO: implement real parser
        var self = this;
        var sheetnum = sheetnum || 0;
        self._sheet[sheetnum] = new Sheet();
        string.replace(Regex.LINEBREAK, Char.NEWLINE).split(Char.NEWLINE).forEach(function (el, i) {
            var row = [];
            el.split(self._delimiter).forEach(function (el) {
                row.push(new Cell(el));
            });
            self._sheet[sheetnum].addRow(row);
        });
        return self;
    };
    CSVParser.prototype.setDelimiter = function (separator) {
        this._delimiter = separator;
        return this;
    };

    var TSVParser = function () {};
    TSVParser.prototype = new CSVParser();
    TSVParser.prototype._delimiter = Char.TAB;
    TSVParser.prototype._filetype = Format.TSV;

    var Parser = {
        CSV : CSVParser,
        TSV : TSVParser
    };

    /////////////
    // Writers
    ////////////

    var BaseWriter = function () {};
    BaseWriter.prototype = {
        addRow      : function (row) {
            throw Exception.UNIMPLEMENTED_METHOD;
        },
        _filetype   : undefined,
        _mimetype   : undefined,
        _sheet      : [],
        insertSheet : function (records) {
            this._sheet.push(records);
        },
        saveFile    : function () {
            // TODO: save to local file
            throw Exception.UNIMPLEMENTED_METHOD
        }
    };

    var CSVWriter = function () {};
    CSVWriter.prototype = new BaseWriter();
    CSVWriter.prototype._delimiter = Char.COMMA;
    CSVWriter.prototype._filetype = Format.CSV;
    CSVWriter.prototype._mimetype = MIMEType.CSV;
    CSVWriter.prototype.saveFile = function () {   
        var self = this;
        var string = '';
        this._sheet[0].forEach(function (el, i) {
            el.forEach(function (el) {
                string += el + self._delimiter;
            });
            string += '\r\n';
        });
        window.open('data:' + this._mimetype + ';base64,' + window.btoa(string));
    };
    CSVWriter.prototype.setDelimiter = function (separator) {
        this._delimiter = separator;
    };

    var TSVWriter = function () {};
    TSVWriter.prototype = new CSVWriter();
    TSVWriter.prototype._delimiter = Char.TAB;
    TSVWriter.prototype._filetype = Format.TSV;
    TSVWriter.prototype._mimetype = MIMEType.TSV;

    var Writer = {
        CSV : CSVWriter,
        TSV : TSVWriter
    };

    /////////////
    // Exports
    ////////////

    var SimpleExcel = {
        Cell                : Cell,
        Exception           : Exception,
        isSupportedBrowser  : Utils.isSupportedBrowser(),
        Parser              : Parser,
        Records             : Records,
        Sheet               : Sheet,
        Writer              : Writer
    };

    window.SimpleExcel = SimpleExcel;

})(this);