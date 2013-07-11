ExileExcel
==========

ExileExcel is an library aim for parse/output Microsoft Excel file with dead easy ways.This project is based on [.NET NPOI 2.0 beta1][0] for read/write excel files.

Usage:
-----------

**Class Define:**

Create your class with ExileAttribute to mapping property with excel file head text.You can define data format explicitily also.

    [ExiliableAttribute("Demo class")]
    class DemoExileData
    {
        public int Id { get; set; }

        public int Id { get; set; }
        [ExileProperty("学号")]
        public string Number { get; set; }
        [ExileProperty("姓名")]
        public string Name { get; set; }
        [ExileProperty("分数","",NPOIDataFormatEnum.NumberInteger)]
        public float Score { get; set; }
        [ExileProperty("考试日期", "YYYY/MM/DD")]
        public DateTime TestDate { get; set; }
    }

**Import:**

Parse from XLS/XLSX files.

    var excelParser = new ExcelParser<DemoExileData>();
    var list= excelParser.Parse(@"d:\1.xlsx");

**Output:**

1.Construct objects form class with ExileAttribute.

    var outData = new List<DemoExileData>
    {
        new DemoExileData
        {
            Id = 1,
            Name = "Alzzl",
            Number = "S0001",
            Score = (float) 50.0,
            TestDate = new DateTime(2013, 1, 3)
        },
        new DemoExileData
        {
            Id = 2,
            Name = "Larvata",
            Number = "S0002",
            Score = (float) 84.5,
            TestDate = new DateTime(2013, 1, 5)
        },
        new DemoExileData
        {
            Id = 3,
            Name = "果子林",
            Number = "S0003",
            Score = (float) 10.5,
            TestDate = new DateTime(2013, 1, 4)
        }
    };


2.Write objects to file

    using (var fs = new FileStream(outFilePath, FileMode.Create, FileAccess.Write))
    {
        extractor.ExcelWriteStream(outData, fs, ExtractType.Excel2007);
    }


Todo:
-----------
1. Output data to word2003/2007.

[0]: https://github.com/tonyqus/npoi