using System;
using System.Collections;
using System.IO;
using HaruP.Common;
using HaruP.Mixins;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace HaruP
{
    public class ExcelExtractor 
    {
        // npoi
        private IWorkbook workbook;
        private ISheet sheet;
        private TemplateMeta templateMeta;

        private ExcelMeta excelMeta;

        public ExcelExtractor()
        {
           excelMeta=new ExcelMeta();
           //excelMeta.Orientation=Orientation.Horizontal;
        }

        private void FillContent(dynamic data)
        {

            if (data is IList || data is IEnumerable)
            {
                for (var i = 0; i < data.Count; i++)
                {
                    WriteSingle(data[i],i);
                }
            }
            else
            {
                WriteSingle(data, 0);
            }
        }

        private void WriteSingle(dynamic singleData, int offset)
        {

            var offsetRow = excelMeta.Orientation == Orientation.Horizontal ? offset : 0;
            var offsetColumn = excelMeta.Orientation == Orientation.Vertical ? offset : 0;

            foreach (var t in templateMeta.Tags)
            {
                var row = (sheet.LastRowNum < t.Value.RowIndex + offsetRow)
                    ? sheet.CreateRow(t.Value.RowIndex + offsetRow)
                    : sheet.GetRow(t.Value.RowIndex + offsetRow);

                var cell = (row.LastCellNum <= t.Value.ColumnIndex + offsetColumn)
                    ? row.CreateCell(t.Value.ColumnIndex + offsetColumn)
                    : row.GetCell(t.Value.ColumnIndex + offsetColumn);
                var val = Utils.GetPropValue(singleData, t.Key);

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

        public void ExcelWriteStream( dynamic data, Stream stream, string templatePath,ExcelMeta excelMeta=null)
        {
            #region Argument check

            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }
            if (string.IsNullOrEmpty(templatePath))
            {
                throw new ArgumentOutOfRangeException("templatePath", "template path not be defined");
            }

            #endregion

            if (excelMeta!=null)
            {
                this.excelMeta = excelMeta;
            }
            
            try
            {
                workbook = WorkbookFactory.Create(templatePath);
            }
            catch (Exception)
            {
            }

            // get tags from template
            ParseTemplateMeta();

            // fill data
            FillContent(data);

            // write to stream
            workbook.Write(stream);
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
