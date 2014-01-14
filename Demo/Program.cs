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
            #region Extract by create new file

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

            const string outFilePath4 = @"out.xls";
            var extractor = new ExileExtractor<DemoExileData>();
            using (var fs = new FileStream(outFilePath4, FileMode.Create))
            {
                extractor.SheetName = "sheet A";
                extractor.TitleText = "this is title";
                extractor.WriteStream(outData, fs, ExileExtractTypes.Excel2003NPOI);
            }
            #endregion

            #region Extract by template
            
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

            const string outFilePathTemp = @"outTemplate.xls";
            var extractorTemp = new ExileExtractor<DemoExileDataTemplate>();
            using (var fs = new FileStream(outFilePathTemp, FileMode.Create))
            {
                extractorTemp.WriteStream(outDataTemp, fs, ExileExtractTypes.Excel2003NPOI, @"template\template.xls");
            }
            #endregion


        }
    }
}
