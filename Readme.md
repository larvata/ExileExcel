ExileExcel
==========

ExileExcel is an library aim for parse/output Microsoft Excel file with dead easy ways.This project is based on [.NET NPOI 2.0 RC][0] for read/write excel files.

Usage:
-----------

**Class Define:**

Create your class with ExileAttribute to mapping property with excel file head text.You can define data format explicitily also.

    [ExileSheetTitle(RowHeight = 40, FontHeight = 40)]
    class DemoExileData:IExilable
    {
        [ExileColumnDataFormat(ColumnType = ExileColumnType.AutoIndex)]
        [ExileHeaderGeneral(HeaderText = "序号", HeaderSequence = 1)]
        public int Id { get; set; }

        [ExileColumnDimension(AutoFit = true,ColumnWidth = 20)]
        [ExileHeaderGeneral(HeaderText = "学号", HeaderSequence = 1)]
        public string Number { get; set; }

        [ExileHeaderGeneral(HeaderText = "姓名", HeaderSequence = 2)]
        public string Name { get; set; }

        [ExileHeaderGeneral(HeaderText = "分数", HeaderSequence = 3)]
        public float Score { get; set; }

        [ExileColumnDataFormat(ColumnCustomDataFormat = "YYYY/MM/DD HH:mm:ss")]
        [ExileHeaderGeneral(HeaderText = "考试日期", HeaderSequence = 4)]
        [ExileColumnDimension(AutoFit = true,ColumnWidth = 20)]
        public DateTime TestDate { get; set; }
    }

**Import:**(not available in current version, checkout `b56f36c` to use it)

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

2.Write data to file

    const string outFilePath = @"out.docx";
    var extractor = new ExileExtractor<DemoExileData>();
   
    using (var fs = new FileStream(outFilePath4, FileMode.Create))
    {
        extractor.SheetName = "sheet A";
        extractor.TitleText = "title";
        extractor.WriteStream(outData, fs, ExileExtractTypes.Word2007OpenXML);
    }

[0]: https://github.com/tonyqus/npoi