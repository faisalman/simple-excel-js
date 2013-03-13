# SimpleExcel.js

Client-side parser & writer for Excel file formats (CSV, XML, XLSX). For server-side solution you might want to check [SimpleExcelPHP](https://github.com/faisalman/simple-excel-php)

## Example

Not yet working, maybe someday...

```html
<!doctype html>
<html>
    <head>
        <script type="text/javascript" src="simple-excel.js"></script>
    </head>
    <body>
        <input type="file" id="fileInput" /><br/>
        <input type="button" id="fileExport" />
        <script type="text/javascript">

            // read a CSV file
            var csvParser = new SimpleExcel.Parser.CSV();
            var fileInput = document.getElementById('fileInput');
            fileInput.addEventListener('change', function (e) {            
                var file = e.target.files[0];
                csvParser.loadFile(file, function () {
                    console.log(csvParser.getSheet()); // print!
                });
            });

            // write an XLSX file
            var xlsxWriter = new SimpleExcel.Writer.XLSX();
            var xlsxSheet = new SimpleExcel.Sheet();
            var Cell = SimpleExcel.Cell;
            xlsxSheet.insertRecord([new Cell('ID', 'TEXT'), new Cell('Nama', 'TEXT']));
            xlsxSheet.insertRecord([new Cell(1, 'NUMBER'), new Cell('Kab. Bogor', 'TEXT']));
            xlsxSheet.insertRecord([new Cell(2, 'NUMBER'), new Cell('Kab. Cianjur', 'TEXT']));
            xlsxSheet.insertRecord([new Cell(3, 'NUMBER'), new Cell('Kab. Sukabumi', 'TEXT']));
            xlsxWriter.insertSheet(xlsxSheet);
            // export when button clicked
            document.getElementById('fileExport').addEventListener('click', function () {            
                xlsxWriter.saveFile(); // pop!
            });

        </script>
    </body>
</html>
```

## License

GPLv2 & MIT License

Copyright © 2013 Faisalman <<fyzlman@gmail.com>>
