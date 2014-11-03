using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using HaruP.Common;
using HaruP.Extractor;
using NPOI.SS.UserModel;

namespace HaruP.Parser
{
    public class ExcelParser:Excel
    {
        
        


        public ExcelParser(string filePath)
        {
//            IWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                this.workbook = WorkbookFactory.Create(file);
            }
//            var worksheet = workbook.GetSheetAt(0);
        }

//        public List<IBaseSheetData> Parse()
//        {
//
//
//        }

        public Sheet GetSheet(int sheetIndex)
        {
            this.sheet = workbook.GetSheetAt(sheetIndex);
            return new Sheet(this.sheet);
        }

        public IEnumerable<Sheet> GetSheets()
        {
            for (var i = 0; i < this.workbook.NumberOfSheets; i++)
            {
                var currentSheet = workbook.GetSheetAt(i);
                yield return new Sheet(currentSheet);
            }
        }
    }
}
