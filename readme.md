# SimpleExcel.js

Client-side parser & writer for Excel file formats (CSV, XML, XLSX)

## Example

Not yet working, maybe someday...

```html
<script type="text/javascript" src="simple-excel.js"></script>
<script type="text/javascript">

    var parser = new SimpleExcel.Parser(SimpleExcel.Format.CSV);
    parser.loadString('Hello, World!');
    parser.getCell(1, 1); // "Hello"

    var writer = new SimpleExcel.Writer(SimpleExcel.Format.XLSX);
    writer.addRow(['Hello', ' World!']);
    writer.addRow(['Hello', ' World!']);
    writer.addRow(['Hello', ' World!']);
    writer.saveFile(); // Pop!

</script>
```

## License

MIT License

Copyright © 2013 Faisalman <<fyzlman@gmail.com>>
