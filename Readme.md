HaruP
==========

HaruP is an library aim for output Microsoft Excel file with dead-easy way.
This project is based on [.NET NPOI][0].


Please chekout branch dotnet20-with-LinqBridge if you are using .NET 2.0

```
git checkout dotnet20-with-LinqBridge
```

Usage:
-----------

## Extractor
build template excel file for output, put mustache.js style interpolate into your cell

#### Template Syntax

Cell value:
```
{{PropertyName}}
```
Cell value with namespace:
```
{{namespace.PropertyName}}
```
Partial cell value:
```
Today is {{PropertyName}
```
Or you can use formula:
```
{{=Formula}}
```
#### Data structure
- HaruP can handle data which implemented IEnumerable.
- Or create it use dynamic type (.net 4.0 only)
- If data with same property name, use namespace

```
var list = new List<dynamic>{
    new{
        Date = DateTime.Now,
        Formula = "A1+B1"},
    new{
        Date = DateTime.Now
        Formula = "A2+B2"}};

var list2 = new List<dynamic>{
    new{
        Date = DateTime.Now},
    new{
        Date = DateTime.Now}};
```

#### Export
```
var extractor = new HaruP.ExcelExtractor(@"template.xls");

extractor.ForkSheet(0, "sheet A")
    // input first list
    .PutData(list)
    // input second list with namespace:S
    .PutData(list2, new SheetMeta{ Namespace = "S"});

extractor.Write("out.xls");
```

## Parser

```
var parser = new ExcelParser(@"forImport.xls");

foreach (var sheets in parser.GetSheets())
{
    var list = sheets.Parse();
}
```


[0]: https://npoi.codeplex.com/
