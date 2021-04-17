Pixcel
======

A utility to convert an Excel file into a bitmap file suitable for importing
into software such as
[DesignaKnit](https://www.softbyte.co.uk/designaknit9.htm).

Setting up your Excel File
--------------------------

Map out your knitting pattern by filling in cells - unfilled will become white
in the bitmap image, filled will become black (unless you specify a color).

Running
-------

Example:

```
pixcel --sheet 'sheetName' inputfile outputfile last_cell
```

(Use `--help`` for details.)

| variable   | description                                                      |
|------------|------------------------------------------------------------------|
| inputfile  | An Excel file (`.xlsx` format)                                   |
| outputfile | A bitmap (`.bmp`) image file                                     |
| last_cell  | The name of the farthest cell from the origin, A1                |
| sheetName  | (optional) The name of the 'sheet' (i.e. tab) in your Excel file |

The `last_cell` must be in standard cell format, such as `A2` or `CF308`.

