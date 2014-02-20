HaruP
==========

HaruP is an library aim for output Microsoft Excel file with dead easy ways.
This project is based on [.NET NPOI 2.0][0].


Usage:
-----------

```
var wotaMixList = new List<dynamic>
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
var e = new ExcelExtractor();
using (var fs = new FileStream("Vertical.xls", FileMode.Create))
{
    e.ExcelWriteStream(wotaMixList, fs, @"template\templateVertical.xls");
}
```
[0]: https://npoi.codeplex.com/
