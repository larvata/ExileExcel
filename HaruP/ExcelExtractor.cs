using System;
using System.Collections;
using System.IO;
using HaruP.Common;
using HaruP.Mixins;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;

namespace HaruP
{
    public class ExcelExtractor 
    {
        // npoi
        private readonly IWorkbook workbook;
        private ISheet sheet;
        private TemplateMeta templateMeta;

        private ExcelMeta excelMeta;

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

           excelMeta = new ExcelMeta();
           //excelMeta.Orientation=Orientation.Horizontal;

           // get tags from template
           ParseTemplateMeta();
        }

        private void WriteSingle(Object singleData, int offset)
        {

            var offsetRow = excelMeta.Orientation == Orientation.Horizontal ? offset : 0;
            var offsetColumn = excelMeta.Orientation == Orientation.Vertical ? offset : 0;

            var isFirst = true;
            foreach (var t in templateMeta.Tags)
            {
                object val;
                try
                {
                    val = Utils.GetPropValue(singleData, t.Key);
                }
                catch (Exception)
                {
                    // property not in current data object 
                    continue;
                }

                var row = isFirst && (offsetRow>0)
                    ? sheet.GetRow(t.Value.RowIndex).CopyRowTo(t.Value.RowIndex + offsetRow)
                    : sheet.GetRow(t.Value.RowIndex+offsetRow);

                var cell = isFirst && (offsetColumn>0)
                    ? row.GetCell(t.Value.ColumnIndex).CopyCellTo(t.Value.ColumnIndex + offsetColumn)
                    : row.GetCell(t.Value.ColumnIndex+offsetColumn);

                isFirst = false;

                // fill value
                if (val == null)
                {
                    cell.SetCellValue(string.Empty);
                }
                else if (val is DateTime)
                {
                    cell.SetCellValue((DateTime) val);
                }
                else if (val is double || val is float || val is int)
                {
                    cell.SetCellValue(Convert.ToDouble(val));
                }
                else
                {
                    cell.SetCellValue(val.ToString());
                }

                // copy cell format to new created
                cell.CellStyle =t.Value.CellStyle;
                
                sheet.SetColumnWidth(cell.ColumnIndex,sheet.GetColumnWidth(t.Value.ColumnIndex));
            }
        }

        public void PutData(IList data, ExcelMeta meta = null)
        {
            this.excelMeta = meta ?? excelMeta;

            // fill data
            for (var i = 0; i < data.Count; i++)
            {
                WriteSingle(data[i], i);
            }
        }

        public void PutData(object data, ExcelMeta meta = null)
        {
            this.excelMeta = meta ?? excelMeta;

            // fill data
            WriteSingle(data, 0);
            
        }

        public void Write(Stream stream)
        {
            #region Argument check

            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            #endregion

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

        private void ParseTemplateMeta()
        {
            templateMeta = new TemplateMeta();
            sheet = workbook.GetSheetAt(0);
            for (var rowNum = 0; rowNum <= sheet.LastRowNum; rowNum++)
            {
                var row = sheet.GetRow(rowNum);

                // each cell in this row was empty 
                if (row == null) continue;

                for (var cellNum = 0; cellNum <= row.LastCellNum; cellNum++)
                {
                    var cell = row.GetCell(cellNum);

                    if (cell == null || cell.CellType != CellType.String) continue;

                    var strVal = cell.StringCellValue;
                    if (!strVal.StartsWith("#{") || !strVal.EndsWith("}")) continue;

                    strVal = strVal.Substring(2).Substring(0, strVal.Length - 3);
                    if (string.IsNullOrEmpty(strVal)) continue;
                    templateMeta.Tags.Add( strVal,cell);
                }
            }

            if (templateMeta.Tags.Count==0)
            {
                throw new InvalidFormatException("cant parse template tags format error. wrap tag with #{}, e.g. #{The Tag}");
            }
        }

    }


}
