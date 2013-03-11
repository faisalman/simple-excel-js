// SimpleExcel.js
// Client-side parser & writer for Excel file formats (CSV, XML, XLSX)
// https://github.com/faisalman/simple-excel-js
// 
// Copyright © 2013 Faisalman <fyzlman@gmail.com>
// Dual licensed under GPLv2 & MIT

(function (global, undefined) {

    'use strict';
    
    ///////////////
    // Constants
    //////////////
    
    var Delimiter = {
        COMMA       : ',',
        SEMICOLON   : ';',
        TAB         : '\t'
    };
    
    var Exception = {    
        CELL_NOT_FOUND              : 'CELL_NOT_FOUND',
        COLUMN_NOT_FOUND            : 'COLUMN_NOT_FOUND',
        ERROR_READING_FILE          : 'ERROR_READING_FILE',
        ERROR_WRITING_FILE          : 'ERROR_WRITING_FILE',
        FIELD_NOT_FOUND             : 'FIELD_NOT_FOUND',
        FILE_NOT_FOUND              : 'FILE_NOT_FOUND',
        FILE_EXTENSION_MISMATCH     : 'FILE_EXTENSION_MISMATCH',
        FILETYPE_NOT_SUPPORTED      : 'FILETYPE_NOT_SUPPORTED',
        INVALID_DOCUMENT_NAMESPACE  : 'INVALID_DOCUMENT_NAMESPACE',
        MALFORMED_JSON              : 'MALFORMED_JSON',
        ROW_NOT_FOUND               : 'ROW_NOT_FOUND',
        UNIMPLEMENTED_METHOD        : 'UNIMPLEMENTED_METHOD',
        UNKNOWN                     : 'UNKNOWN'
    };
    
    var Format = {        
        CSV     : 'CSV',
        XLSX    : 'XLSX',
        XML     : 'XML'
    };

    /////////////
    // Parsers
    ////////////

    var BaseParser = function () {};
    BaseParser.prototype = {
        columnLength: undefined,
        extension   : undefined,
        field       : {},
        getCell     : function (colNum, rowNum) {
            throw Exception.UNIMPLEMENTED_METHOD;            
        },
        getColumn   : function (colNum) {
            throw Exception.UNIMPLEMENTED_METHOD;            
        },
        getField    : function () {
            return this.field;
        },
        getRow      : function (rowNum) {
            throw Exception.UNIMPLEMENTED_METHOD;
        },
        loadFile    : function (filePath) {
            throw Exception.UNIMPLEMENTED_METHOD;
        },
        loadString  : function (string) {
            throw Exception.UNIMPLEMENTED_METHOD;
        },
        rowLength   : undefined
    };
    
    var CSVParser = function () {
        this.delimiter = Delimiter.COMMA;
    };
    CSVParser.prototype = new BaseParser();
    CSVParser.prototype.extension = Format.CSV;
    CSVParser.prototype.loadFile = function (filePath) {
        var isFileFound;
        if (isFileFound) {
            // TODO: try open file
            var isFileReadable;
            if (isFileReadable) {
                // TODO: read the content
                var content;
                this.loadString(content);
            } else {
                throw Exception.ERROR_READING_FILE;
            }
        } else {
            throw Exception.FILE_NOT_FOUND;
        }
    };
    CSVParser.prototype.loadString = function (string) {
        // TODO: parse string
        this.field = string;
    };
    CSVParser.prototype.setDelimiter = function (separator) {
        this.delimiter = separator;
    };
    
    var Parsers = {
        CSV : CSVParser
    };
    
    var Parser = function (format) {
        if (Parsers.hasOwnProperty(format)) {
            return new Parsers[format]();
        } else {
            throw Exception.FILETYPE_NOT_SUPPORTED;
        }
    };
    
    /////////////
    // Writers
    ////////////
    
    var BaseWriter = function () {};
    BaseWriter.prototype = {
        addRow      : function (row) {
            throw Exception.UNIMPLEMENTED_METHOD;
        },
        columnLength: undefined,
        extension   : undefined,
        field       : {},
        saveFile    : function () {
            throw Exception.UNIMPLEMENTED_METHOD;
        },
        saveString  : function () {
            throw Exception.UNIMPLEMENTED_METHOD;
        },
        setData     : function (data) {
            throw Exception.UNIMPLEMENTED_METHOD;        
        },
        rowLength   : undefined
    };
    
    var CSVWriter = function () {
        this.delimiter = Delimiter.COMMA;
    };
    CSVWriter.prototype = new BaseWriter();
    CSVWriter.prototype.extension = Format.CSV;
    CSVWriter.prototype.setDelimiter = function (separator) {
        this.delimiter = separator;
    };
    
    var Writers = {
        CSV : CSVWriter
    };
    
    var Writer = function (format) {
        if (Writers.hasOwnProperty(format)) {
            return new Writers[format]();
        } else {
            throw Exception.FILETYPE_NOT_SUPPORTED;
        }
    };
    
    /////////////
    // Exports
    ////////////
    
    var SimpleExcel = {
        Delimiter : Delimiter,
        Exception : Exception,
        Format : Format,
        Parser : Parser,
        Writer : Writer
    };
    
    global.SimpleExcel = SimpleExcel;

})(this);


///
// // Quick test
//
// var parser = new SimpleExcel.Parser(SimpleExcel.Format.CSV);
// parser.loadString('Hello, World!');
// console.log(parser.getField());
