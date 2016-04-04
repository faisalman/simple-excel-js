// SimpleExcel.js v0.1.3
// Client-side script to easily parse / convert / write any Microsoft Excel XLSX / XML / CSV / TSV / HTML / JSON / etc formats
// https://github.com/faisalman/simple-excel-js
// 
// Copyright © 2013-2014 Faisal Salman <fyzlman@gmail.com>
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
        ROW_NOT_FOUND               : 'ROW_NOT_FOUND',
        ERROR_READING_FILE          : 'ERROR_READING_FILE',
        ERROR_WRITING_FILE          : 'ERROR_WRITING_FILE',
        FILE_NOT_FOUND              : 'FILE_NOT_FOUND',
        //FILE_EXTENSION_MISMATCH     : 'FILE_EXTENSION_MISMATCH',
        FILETYPE_NOT_SUPPORTED      : 'FILETYPE_NOT_SUPPORTED',
        INVALID_DOCUMENT_FORMAT     : 'INVALID_DOCUMENT_FORMAT',
        INVALID_DOCUMENT_NAMESPACE  : 'INVALID_DOCUMENT_NAMESPACE',
        MALFORMED_JSON              : 'MALFORMED_JSON',
        UNIMPLEMENTED_METHOD        : 'UNIMPLEMENTED_METHOD',
        UNKNOWN_ERROR               : 'UNKNOWN_ERROR',
        UNSUPPORTED_BROWSER         : 'UNSUPPORTED_BROWSER'
    };

    var Format = {        
        CSV     : 'csv',
        HTML    : 'html',
        JSON    : 'json',
        TSV     : 'tsv',
        XLS     : 'xls',
        XLSX    : 'xlsx',
        XML     : 'xml'
    };

    var MIMEType = {
        CSV     : 'text/csv',
        HTML    : 'text/html',
        JSON    : 'application/json',
        TSV     : 'text/tab-separated-values',
        XLS     : 'application/vnd.ms-excel',
        XLSX    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        XML     : 'text/xml',
        XML2003 : 'application/xml'
    };

    var Regex = {
        FILENAME    : /.*\./g,
        LINEBREAK   : /\r\n?|\n/g,
        COMMA       : /(,)(?=(?:[^"]|"[^"]*")*$)/g,
        QUOTATION   : /(^")(.*)("$)/g,
        TWO_QUOTES  : /""/g
    };

    var Utils = {
        getFiletype : function (filename) {
            return filename.replace(Regex.FILENAME, '');
        },
        isEqual     : function (str1, str2, ignoreCase) {
            return ignoreCase ? str1.toLowerCase() == str2.toLowerCase() : str1 == str2;
        },
        isSupportedBrowser: function() {
            return !![].forEach && !!window.FileReader;
        },
        overrideProperties : function (old, fresh) {
            for (var i in old) {
                if (old.hasOwnProperty(i)) {
                    old[i] = fresh.hasOwnProperty(i) ? fresh[i] : old[i];
                }
            }
            return old;
        }
    };
    
    /////////////////////////////
    // Spreadsheet Constructors
    ////////////////////////////

    var Cell = function (value, dataType) {
        var defaults = {
            value    : value || '',
            dataType : dataType || DataType.TEXT
        };
        if (typeof value == typeof {}) {
            defaults = Utils.overrideProperties(defaults, value);
        }
        this.value = defaults.value;
        this.dataType = defaults.dataType;
        this.toString = function () {
            return value.toString();
        };
    };
        
    var Records = function() {};
    Records.prototype = [];
    Records.prototype.getCell = function(colNum, rowNum) {
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
    
    var Sheet = function () {
        this.records = new Records();
    };
    Sheet.prototype.getCell = function (colNum, rowNum) {
        return this.records.getCell(colNum, rowNum);
    };
    Sheet.prototype.getColumn = function (colNum) {
        return this.records.getColumn(colNum);
    };
    Sheet.prototype.getRow = function (rowNum) {
        return this.records.getRow(rowNum);
    };
    Sheet.prototype.insertRecord = function (array) {
        this.records.push(array);
        return this;
    };
    Sheet.prototype.removeRecord = function (index) {
        this.records.splice(index - 1, 1);
        return this;
    };
    Sheet.prototype.setRecords = function (records) {
        this.records = records;
        return this;
    };
    
    /////////////
    // Parsers
    ////////////

    // Base Class
    var BaseParser = function () {};
    BaseParser.prototype = {
        _filetype   : '',
        _sheet      : [],
        getSheet    : function(number) {
            number = number || 1;
            return this._sheet[number - 1].records;
        },
        loadFile    : function (file, callback) {
            var self = this;
            //var filetype = Utils.getFiletype(file.name);
            //if (Utils.isEqual(filetype, self._filetype, true)) {
                var reader = new FileReader();
                reader.onload = function () {
                    self.loadString(this.result, 0);
                    callback.apply(self);
                };
                reader.readAsText(file);
            //} else {
                //throw Exception.FILE_EXTENSION_MISMATCH;
            //}
            return self;
        },
        loadString  : function (string, sheetnum) {
            throw Exception.UNIMPLEMENTED_METHOD;
        }
    };

    // CSV
    var CSVParser = function () {};
    CSVParser.prototype = new BaseParser();
    CSVParser.prototype._delimiter = Char.COMMA;
    CSVParser.prototype._filetype = Format.CSV;
    CSVParser.prototype.loadString = function (str, sheetnum) {
        // TODO: implement real CSV parser
        var self = this;
        sheetnum = sheetnum || 0;
        self._sheet[sheetnum] = new Sheet();       
        
        str.replace(Regex.LINEBREAK, Char.NEWLINE)
           .split(Char.NEWLINE)
           .forEach(function(el, i)
        {
            var sp = el.split(Regex.COMMA);
            var row = [];
            sp.forEach(function(cellText) {
                if (cellText !== self._delimiter) {
                    cellText = cellText.replace(Regex.QUOTATION, "$2");
                    cellText = cellText.replace(Regex.TWO_QUOTES, "\"");
                    row.push(new Cell(cellText));
                }
            });
            self._sheet[sheetnum].insertRecord(row);
        });
        return self;
    };
    CSVParser.prototype.setDelimiter = function (separator) {
        this._delimiter = separator;
        return this;
    };
    
    // HTML
    var HTMLParser = function () {};
    HTMLParser.prototype = new BaseParser();
    HTMLParser.prototype._filetype = Format.HTML;
    HTMLParser.prototype.loadString = function(str, sheetnum) {
        var self = this;
        var domParser = new DOMParser();
        var domTree = domParser.parseFromString(str, MIMEType.HTML);
        var sheets = domTree.getElementsByTagName('table');
        sheetnum = sheetnum || 0;
        [].forEach.call(sheets, function(el, i) {
            self._sheet[sheetnum] = new Sheet();
            var rows = el.getElementsByTagName('tr');
            [].forEach.call(rows, function (el, i) {
                var cells = el.getElementsByTagName('td');
                var row = [];
                [].forEach.call(cells, function (el, i) {
                    row.push(new Cell(el.innerHTML));
                });
                self._sheet[sheetnum].insertRecord(row);
            });
            sheetnum++;
        });
        return self;
    };

    // TSV
    var TSVParser = function () {};
    TSVParser.prototype = new CSVParser();
    TSVParser.prototype._delimiter = Char.TAB;
    TSVParser.prototype._filetype = Format.TSV;

    // XML
    var XMLParser = function () {};
    XMLParser.prototype = new BaseParser();
    XMLParser.prototype._filetype = Format.XML;
    XMLParser.prototype.loadString = function(str, sheetnum) {
        var self = this;
        var domParser = new DOMParser();
        var domTree = domParser.parseFromString(str, MIMEType.XML);
        var sheets = domTree.getElementsByTagName('Worksheet');
        sheetnum = sheetnum || 0;
        [].forEach.call(sheets, function(el, i) {
            self._sheet[sheetnum] = new Sheet();
            var rows = el.getElementsByTagName('Row');
            [].forEach.call(rows, function (el, i) {
                var cells = el.getElementsByTagName('Data');
                var row = [];
                [].forEach.call(cells, function (el, i) {
                    row.push(new Cell(el.innerHTML));
                });
                self._sheet[sheetnum].insertRecord(row);
            });
            sheetnum++;
        }); 
        return self;
    };

    // Export var
    var Parser = {
        CSV : CSVParser,
        HTML: HTMLParser,
        TSV : TSVParser,
        XML : XMLParser
    };

    /////////////
    // Writers
    ////////////

    // Base Class
    var BaseWriter = function () {};
    BaseWriter.prototype = {
        _filetype   : '',
        _mimetype   : '',
        _sheet      : [],
        getSheet    : function(number) {
            number = number || 1;
            return this._sheet[number - 1].records;
        },
        getString   : function () {
            throw Exception.UNIMPLEMENTED_METHOD;
        },
        insertSheet : function (data) {
            if (!!data.records) {
                this._sheet.push(data);
            } else {
                var sheet = new Sheet();
                sheet.setRecords(data);
                this._sheet.push(sheet);
            }
            return this;
        },
        removeSheet : function (index) {
            this._sheet.splice(index - 1, 1);
            return this;
        },
        saveFile    : function () {
            // TODO: find a reliable way to save as local file
            window.open('data:' + this._mimetype + ';base64,' + window.btoa(this.getString()));            
            return this;
        }
    };

    // CSV
    var CSVWriter = function () {};
    CSVWriter.prototype = new BaseWriter();
    CSVWriter.prototype._delimiter = Char.COMMA;
    CSVWriter.prototype._filetype = Format.CSV;
    CSVWriter.prototype._mimetype = MIMEType.CSV;
    CSVWriter.prototype.getString = function () {
        // TODO: implement real CSV writer
        var self = this;
        var string = '';
        this.getSheet(1).forEach(function (el, i) {
            el.forEach(function (el) {
                string += el + self._delimiter;
            });
            string += '\r\n';
        });
        return string;
    };
    CSVWriter.prototype.setDelimiter = function (separator) {
        this._delimiter = separator;
        return this;
    };

    // TSV
    var TSVWriter = function () {};
    TSVWriter.prototype = new CSVWriter();
    TSVWriter.prototype._delimiter = Char.TAB;
    TSVWriter.prototype._filetype = Format.TSV;
    TSVWriter.prototype._mimetype = MIMEType.TSV;
    
    // Export var
    var Writer = {
        CSV : CSVWriter,
        TSV : TSVWriter
    };

    /////////////
    // Exports
    ////////////

    var SimpleExcel = {
        Cell                : Cell,
        DataType            : DataType,
        Exception           : Exception,
        isSupportedBrowser  : Utils.isSupportedBrowser(),
        Parser              : Parser,
        Sheet               : Sheet,
        Writer              : Writer
    };

    window.SimpleExcel = SimpleExcel;

})(this);
