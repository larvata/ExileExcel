using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExileExcel;
using ExileExcel.Common;
using System.Web;

namespace Demo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //ExileExcel e=new ExileExcel();
            //var exileParser = new ExileParser();
            var excelParser = new ExcelParser<DemoExileData>();

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


            using (var fs = new FileStream(outFilePath, FileMode.Create, FileAccess.Write))
            {
                excelParser.ExcelOutputStream(outData, fs);
            }






            //var list = ExileExcel.Common.ReflectiveEnumerator.GetTypesWithAttribute(typeof(ExiliableAttribute)).ToList();
            //ExileExcel.ParserUtilty.GetNameAttributePair(list);
            //parser.SaveExcelFile();
        }
    }
}
