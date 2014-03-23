HaruP
==========

HaruP is an library aim for output Microsoft Excel file with dead easy ways.
This project is based on [.NET NPOI 2.0][0].


Usage:
-----------

```
var list = new List<dynamic>
{
    new
    {
        Tiger = 1,
        Fire = "Fire1",
        Cyber = DateTime.Now,
        Fiber = (double) 2,
        Diver = (Single) 3,
        Viber = "Viber1"
    },
    new
    {
        Tiger = 4,
        Cyber = DateTime.Now.AddDays(1),
        Fiber = (double) 5,
        Diver = (Single) 6,
        Viber = "Viber2"
    }
};

var stat = new
{
    Count = "Count: " + list.Count;
};

var extractor = new ExcelExtractor(@"template.xls");

extractor.ForkSheet(0, "sheet A")
    .PutData(list)
    .PutData(stat)

extractor.Write("out.xls");

```
[0]: https://npoi.codeplex.com/
