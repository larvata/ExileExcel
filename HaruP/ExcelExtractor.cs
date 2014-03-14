using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text.RegularExpressions;

using HaruP.Mixins;
using NPOI.HSSF.Record.PivotTable;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;

namespace HaruP
{
    public class ExcelExtractor:Excel
    {

        public ExcelExtractor(string templatePath)
        {
            #region Argument check

            if (string.IsNullOrEmpty(templatePath))
            {
                throw new ArgumentOutOfRangeException("templatePath", "template path not be defined");
            }

            #endregion

           try
           {
               workbook = WorkbookFactory.Create(templatePath);
           }
           catch (WorkbookNotFoundException ex)
           {

           }

        }

        public Sheet ForkSheet(int sheetIndex, string forkedSheetName)
        {
            workbook.SetSheetHidden(sheetIndex,1);

            this.sheet = workbook.CloneSheet(sheetIndex);
            var index = this.workbook.GetSheetIndex(this.sheet);
            this.workbook.SetSheetName(index, forkedSheetName);
            return new Sheet(this.sheet);
        }

        public Sheet GetSheet(int sheetIndex)
        {
            this.sheet = workbook.GetSheetAt(sheetIndex);
            return new Sheet(this.sheet);
        }

        public void Write(Stream stream)
        {
            #region Argument check

            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            #endregion

            // recalculate formula
            sheet.ForceFormulaRecalculation = true;

            // write to stream
            workbook.Write(stream);
        }

        public void Write(string filePath)
        {
            using (var fs=new FileStream(filePath, FileMode.Create))
            {
                this.Write(fs);
            }
        }

    }

}
