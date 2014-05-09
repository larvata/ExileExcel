HaruP
==========

HaruP is an library aim for output Microsoft Excel file with dead easy ways.
This project is based on [.NET NPOI 2.0][0].


Usage:
-----------

build template excel file for output, put mustache.js style interpolate into your cell

#### Template Syntax
```
// cell value
{{PropertyName}}

// partial cell value
Today is {{PropertyName}

// Or you can use formula
{{=Formula}}
```
#### Data structure
- HaruP can handle data implement IEnumerable.
- Or create it use dynamic type (.net 4.0 only)
- If datas with same property name, use namespace

```
var list = new List<dynamic>
{
    new
    {
        Id = 1,
        Date = DateTime.Now,
        Formula = "A1+B1"
    },
    new
    {
        Id = 2,
        Date = DateTime.Now
        Formula = "A2+B2"
    }
};

var list2 = new List<dynamic>
{
    new
    {
        Id = 1,
        Date = DateTime.Now,
    },
    new
    {
        Id = 2,
        Date = DateTime.Now
    }
};

// namespace define
var meta = new SheetMeta
{
    Namespace = "S",
};
```

#### Export
```
var extractor = new HaruP.ExcelExtractor(@"template.xls");

extractor.ForkSheet(0, "sheet A")
    // input first list
    .PutData(list)
    // input second list with namespace:S
    .PutData(list2,meta);

extractor.Write("out.xls");
```

[0]: https://npoi.codeplex.com/
