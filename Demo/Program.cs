using System;
using System.Collections.Generic;
using System.IO;
using ExileExcel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Demo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            #region Extract
            
            var outData = new List<DemoExileDataTemplate>
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

            const string outFilePath4 = @"d:\out_etmplate.xls";
            //const string outFilePath3 = @"d:\out3.xlsx";
            var extractor = new ExileExtractor<DemoExileDataTemplate>();
            using (var fs = new FileStream(outFilePath4, FileMode.Create))
            {
                extractor.SheetName = "sheet A";
                extractor.TitleText = "titttttttttttttttttile";
                extractor.WriteStream(outData, fs, ExileExtractTypes.Excel2003NPOI, @"d:\template.xls");
            }
            #endregion


        }
    }
}
