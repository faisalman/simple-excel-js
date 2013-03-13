# SimpleExcel.js

Client-side parser & writer for Excel file formats (CSV, XML, XLSX)

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
            xlsxSheet.addRow([new Cell('ID'), new Cell('Nama']));
            xlsxSheet.addRow([new Cell('1'), new Cell('Kab. Bogor']));
            xlsxSheet.addRow([new Cell('2'), new Cell('Kab. Cianjur']));
            xlsxSheet.addRow([new Cell('3'), new Cell('Kab. Sukabumi']));
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

MIT License

Copyright © 2013 Faisalman <<fyzlman@gmail.com>>
