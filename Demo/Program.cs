using ExileExcel;
using ExileExcel.Common;
using System;
using System.Collections.Generic;
using System.IO;

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

            var excelParser = new ExileParser<DemoExileData>();
            var list = excelParser.Parse(@"d:\1.xlsx");

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
                    Id = 1,
                    Name = "果子林",
                    Number = "S0003",
                    Score = (float) 10.5,
                    TestDate = new DateTime(2013, 1, 4)
                }
            };


            const string outFilePath = @"d:\out.xlsx";
            var extractor = new ExileExtractor<DemoExileData>();
           

            using (var fs = new FileStream(outFilePath, FileMode.Create, FileAccess.Write))
            {
                extractor.ExcelWriteStream(outData, fs, ExtractType.Excel2007);
            }

        }
    }
}
