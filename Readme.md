ExileExcel
==========

ExileExcel is an library aim for parse/output Microsoft Excel file with dead easy ways.This project is based on [.NET NPOI 2.0 RC][0] for read/write excel files.

Features:
 - Export data to xls,xlsx,docx
 - Export with template

Usage:
-----------

**Class Define:**

Create your class with ExileAttribute to mapping property with excel file head text.You can define data format explicitily also.

    // define data area
    [ExileSheetDataArea(StartRowNum = 2)]
    // hide data header which is already in template file
    [ExileSheetTitle(HideHeader=true )]
    public class DemoExileDataTemplate : IExilable
    {
        // first row auto index
        [ExileColumnDataFormat(ColumnType = ExileColumnType.AutoIndex)]
        [ExileHeaderGeneral(HeaderSequence = 1)]
        public int Id { get; set; }

        // set width of column and autofit text
        [ExileColumnDimension(AutoFit = true, ColumnWidth = 20)]
        [ExileHeaderGeneral(HeaderSequence = 2)]
        public string Number { get; set; }

        [ExileHeaderGeneral(HeaderSequence = 3)]
        public string Name { get; set; }

        [ExileHeaderGeneral(HeaderSequence = 4)]
        public float Score { get; set; }
        
        // define cell data format
        [ExileColumnDataFormat(ColumnCustomDataFormat = "YYYY/MM/DD HH:mm:ss")]
        [ExileHeaderGeneral(HeaderSequence = 5)]
        public DateTime TestDate { get; set; }
    }

**Import:**(not available in current version, checkout `b56f36c` to use it)

Parse from XLS/XLSX files.

    var excelParser = new ExcelParser<DemoExileData>();
    var list= excelParser.Parse(@"d:\1.xlsx");

**Output:**

1.Construct objects form class with ExileAttribute.

    var outDataTemp = new List<DemoExileDataTemplate>
    {
        new DemoExileDataTemplate
        {
            Id = 1,
            Name = "Alzzl",
            Number = "S0001",
            Score = (float) 50.0,
            TestDate = new DateTime(2013, 1, 3)
        },
        new DemoExileDataTemplate
        {
            Id = 1,
            Name = "Larvata",
            Number = "S0002",
            Score = (float) 84.5,
            TestDate = new DateTime(2013, 1, 5)
        },
        new DemoExileDataTemplate
        {
            Id=1,
            Name = "LALALA",
            Number = "1",
            Score = (float)100.0,
            TestDate = new DateTime(2012,5,5)
        },
        new DemoExileDataTemplate
        {
            Id = 1,
            Name = "果子林",
            Number = "S0003",
            Score = (float) 10.5,
            TestDate = new DateTime(2013, 1, 4)
        },


    };

2.Write data with template

    const string outFilePathTemp = @"outTemplate.xls";
    var extractorTemp = new ExileExtractor<DemoExileDataTemplate>();
    using (var fs = new FileStream(outFilePathTemp, FileMode.Create))
    {
        extractorTemp.WriteStream(outDataTemp, fs, ExileExtractTypes.Excel2003NPOI, @"template\template.xls");
    }

[0]: https://npoi.codeplex.com/