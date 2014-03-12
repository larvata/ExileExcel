using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text.RegularExpressions;
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

        }

        private void WriteSingle(Object singleData, int offset)
        {

            var offsetRow = excelMeta.Orientation == Orientation.Horizontal ? offset : 0;
            var offsetColumn = excelMeta.Orientation == Orientation.Vertical ? offset : 0;

            var isFirst = true;
            var matchedTags = string.IsNullOrEmpty(excelMeta.Namespace)
                ? excelMeta.Tags.Where(t => !t.TagId.Contains(".")).ToList()
                : excelMeta.Tags.Where(t => t.TagId.StartsWith(excelMeta.Namespace)).ToList();


            foreach (var t in matchedTags)
            {
                object cellValue;

                try
                {
                    cellValue = Utils.GetPropValue(singleData, t.TagId.Remove(0, excelMeta.Namespace.Length));
                }
                catch (Exception)
                {
                    // property not in current data object 
                    continue;
                }

                var row=isFirst && (offsetRow > 0)
                    ?sheet.GetRow(t.Cell.RowIndex).CopyRowTo(t.Cell.RowIndex + offsetRow)
                    :sheet.GetRow(t.Cell.RowIndex + offsetRow);


                var cell = isFirst && (offsetColumn > 0)
                    ? row.GetCell(t.Cell.ColumnIndex).CopyCellTo(t.Cell.ColumnIndex + offsetColumn)
                    : row.GetCell(t.Cell.ColumnIndex + offsetColumn);

                isFirst = false;

                // fill value
                if (cellValue == null)
                {
                    cell.SetCellValue(string.Empty);
                }
                else if (cellValue is DateTime)
                {
                    cell.SetCellValue((DateTime) cellValue);
                }
                else if (cellValue is double || cellValue is float || cellValue is int)
                {
                    cell.SetCellValue(Convert.ToDouble(cellValue));
                }
                else
                {
                    cell.SetCellValue(cellValue.ToString());
                }

                // copy cell format to new created
                cell.CellStyle = t.Cell.CellStyle;

                sheet.SetColumnWidth(cell.ColumnIndex, sheet.GetColumnWidth(t.Cell.ColumnIndex));
            }
        }


        public void PutData(IList data)
        {
            // get tags from template
            ParseTemplateMeta(this.excelMeta.SheetIndex);

            // fill data
            for (var i = 0; i < data.Count; i++)
            {
                WriteSingle(data[i], i);
            }

            // reset namespace
            this.excelMeta.Namespace = string.Empty;
        }

        public void PutData(IList data, ExcelMeta meta)
        {
            this.excelMeta = meta ?? excelMeta;
            PutData(data);

        }

        public void PutData(IList data, int sheetIndex)
        {
            this.excelMeta.SheetIndex=sheetIndex;
            PutData(data);
        }

        public void PutData(IList data, string tagNamespace)
        {
            this.excelMeta.Namespace = tagNamespace;
            PutData(data);
        }

        public void PutData(object data)
        {
            // get tags from template
            ParseTemplateMeta(this.excelMeta.SheetIndex);

            // fill data
            WriteSingle(data, 0);

            // reset namespace
            this.excelMeta.Namespace = string.Empty;
        }

        public void PutData(object data, ExcelMeta meta)
        {
            this.excelMeta = meta ?? excelMeta;
            // fill data
            PutData(data);
        }

        public void PutData(object data, int sheetIndex)
        {
            this.excelMeta.SheetIndex = sheetIndex;
            // fill data
            PutData(data);
        }

        public void PutData(object data, string tagNamespace)
        {
            this.excelMeta.Namespace = tagNamespace;
            // fill data
            PutData(data);
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

        private void ParseTemplateMeta(int sheetIndex)
        {
            sheet = workbook.GetSheetAt(sheetIndex);
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

                    var tagId = Regex.Match(strVal, @"#{([\w\.]+)}").Groups[1].Value;

                    if (string.IsNullOrEmpty(tagId)) continue;

                    excelMeta.Tags.Add(new TagMeta
                    {
                        Cell = cell,
                        TagId = tagId,
                        TemplateText = strVal
                    });

                }
            }
        }

    }


}
