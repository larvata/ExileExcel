using System;
using System.Collections.Generic;
using System.IO;
using ExileNPOI;
using ExileNPOI.Common;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Demo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
//            XWPFDocument doc = null;
//            
//            using (var fs=new FileStream(@"d:\1.doc",FileMode.Open,FileAccess.Read))
//            {
//                doc= new XWPFDocument(fs);
//                    
//            }
//            
//
//            return;

//            var excelParser = new ExileParser<DemoExileData>();
//            var list = excelParser.Parse(@"d:\1.xlsx");

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
                    Id = 1,
                    Name = "Larvata",
                    Number = "S0002",
                    Score = (float) 84.5,
                    TestDate = new DateTime(2013, 1, 5)
                },
                new DemoExileData
                {
                    Id=1,
                    Name = "LALALA",
                    Number = "1",
                    Score = (float)100.0,
                    TestDate = new DateTime(2012,5,5)
                },
                new DemoExileData
                {
                    Id = 1,
                    Name = "果子林",
                    Number = "S0003",
                    Score = (float) 10.5,
                    TestDate = new DateTime(2013, 1, 4)
                },


            };
//
//
//            const string outFilePath = @"d:\out.xlsx";
//            var extractor = new ExileExtractor<DemoExileData>("sheet","title");
//           
//            using (var fs = new FileStream(outFilePath, FileMode.Create, FileAccess.Write))
//            {
//                extractor.ExcelWriteStream(outData, fs, ExcelExtractTypes.Excel2007);
//            }
//
//
//            /////////////////////////////
//            IWorkbook workbook=new XSSFWorkbook();
//            var sheet= workbook.CreateSheet("sheet A");
//            sheet.Compose(outData,"Big Title");
//            const string outFilePath2 = @"d:\out2.xlsx";
//            using (var fs=new FileStream(outFilePath2,FileMode.Create,FileAccess.Write))
//            {
//                workbook.Write(fs);
//            }

            const string outFilePath4 = @"d:\out4.docx";
            const string outFilePath3 = @"d:\out3.xlsx";
            var extractor = new ExileExtractor<DemoExileData>();
            using (var fs = new FileStream(outFilePath4, FileMode.Create))
            {
                extractor.SheetName = "sheet A";
                extractor.TitleText = "titttttttttttttttttile";
                extractor.WriteStream(outData,fs,ExileExtractTypes.Word2007);
            }

//            var t=new TestForTemplate();
//            t.BuildData();
        }
    }
}
