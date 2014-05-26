using System;
using System.Collections;
using System.Linq;
using System.Text.RegularExpressions;
using HaruP.Common;
using NPOI.SS.UserModel;

namespace HaruP.Extractor
{ 
    public class Sheet:Excel
    {
        private const string Interpolate = @"{{([\w\.]+)}}";
        private const string Formula = @"{{=([\w\.]+)}}";

        public Sheet(ISheet sheet)
        {
            this.sheet = sheet;
            this.workbook = sheet.Workbook;
            this.sheetMeta = new SheetMeta();
        }

        public void Delete()
        {
            var index = this.workbook.GetSheetIndex(this.sheet);
            this.workbook.RemoveSheetAt(index);
        }

        public Sheet PutData(object data)
        {
            // get tags from template
            ParseTemplateMeta();

            // determine type of [object data]
            if (data is IEnumerable)
            {
                // cast to list for get length
                var list=(data as IEnumerable).Cast<Object>().ToArray();

                for (var i = 0; i < list.Length; i++)
                {
                    WriteSingle(list[i], i,(i==list.Length-1));
                }
            }
            else
            {
                // fill data
                WriteSingle(data, 0);
            }
            
            return this;
        }

        public Sheet PutData(object data, SheetMeta meta)
        {
            this.sheetMeta = meta ?? this.sheetMeta;
            // fill data
            return PutData(data);
        }

        private void ParseTemplateMeta()
        {
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

                    var matchText = Regex.Match(strVal, Interpolate);
                    while (matchText.Success)
                    {
                        sheetMeta.Tags.Add(new TagMeta
                        {
                            Cell = cell,
                            TagId = matchText.Groups[1].ToString(),
                            TemplateText = matchText.Groups[0].ToString(),
                            TagType = TagType.Text
                        });    
                        matchText = matchText.NextMatch();
                    }

                    var matchFormula = Regex.Match(strVal, Formula);
                    while (matchFormula.Success)
                    {
                        sheetMeta.Tags.Add(new TagMeta
                        {
                            Cell = cell,
                            TagId = matchFormula.Groups[1].ToString(),
                            TemplateText = matchFormula.Groups[0].ToString(),
                            TagType = TagType.Formula
                        });
                        matchFormula = matchFormula.NextMatch();
                    }

                }
            }
        }

        private void WriteSingle(Object singleData, int offset,bool isLastObject=true)
        {
            var offsetRow = sheetMeta.Orientation == Orientation.Horizontal ? offset : 0;
            var offsetColumn = sheetMeta.Orientation == Orientation.Vertical ? offset : 0;

            var isFirstCell = true;
            var matchedTags = string.IsNullOrEmpty(sheetMeta.Namespace)
                ? sheetMeta.Tags.Where(t => !t.TagId.Contains(".")).ToList()
                : sheetMeta.Tags.Where(t => t.TagId.StartsWith(sheetMeta.Namespace)).ToList();

            foreach (var t in matchedTags)
            {
                object cellValue;

                try
                {
                    cellValue = Utils.GetPropValue(singleData, t.TagId.Remove(0, sheetMeta.Namespace.Length));
                }
                catch (Exception)
                {
                    // property not in current data object 
                    continue;
                }

                IRow row;
                if (isFirstCell && !isLastObject)
                {
                    row = sheet.GetRow(t.Cell.RowIndex + offsetRow)
                        .CopyRowToAdvance(t.Cell.RowIndex + offsetRow + 1, sheetMeta.RowHeight)
                        .CellFormulaShift(1);
                }
                else
                {
                    row = sheet.GetRow(t.Cell.RowIndex + offsetRow);
                }

                // todo check code
                var cell = isFirstCell && (offsetColumn > 0)
                    ? row.GetCell(t.Cell.ColumnIndex).CopyCellTo(t.Cell.ColumnIndex + offsetColumn)
                    : row.GetCell(t.Cell.ColumnIndex + offsetColumn);

                isFirstCell = false;

                switch (t.TagType)
                {
                    case TagType.Text:
                        FillCellValue(cellValue, cell,t.TemplateText);
                        break;
                    case TagType.Formula:
                        FillCellFormula(cellValue, cell);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                // copy cell format to new created
                cell.CellStyle = t.Cell.CellStyle;
                sheet.SetColumnWidth(cell.ColumnIndex, sheet.GetColumnWidth(t.Cell.ColumnIndex));
            }
        }

        private static void FillCellFormula(object cellValue, ICell cell)
        {
            if (cellValue==null)
            {
                cell.SetCellValue(string.Empty);
            }
            else
            {
                cell.CellFormula = cellValue.ToString();
            }
        }

        private static void FillCellValue(object cellValue, ICell cell,string templateText)
        {
            // fill value
            var currentCellValue = cell.StringCellValue;
            cell.SetCellValue(currentCellValue.Replace(templateText,(cellValue??string.Empty).ToString()));
        }
    }
}
