using System;
using System.Collections.Generic;
using ExileExcel.Attribute;
using ExileExcel.Common;
using ExileExcel.Interface;
using ExileExcel.Mixins;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExileExcel
{
    internal class ExileExcelExtractor<T> : IExtractor<T> where T:IExilable
    {
        internal ExileDocumentMeta DocumentMeta { get; set; }
        private readonly Dictionary<string, ICellStyle> _cellFormatCache;

        private readonly Dictionary<string, ExileColumnDataFormatAttribute> _dataFormatAttr;
        private readonly ISheet _sheet;

        internal ExileExcelExtractor(ISheet sheet)
        {
            // fill all buildin dataformat for avoid re-declare same format
            _cellFormatCache = new Dictionary<string, ICellStyle>();

            // init attribute cache
            _dataFormatAttr = new Dictionary<string, ExileColumnDataFormatAttribute>();
            DocumentMeta = Utils.GetTypeMatched<T>();


            _sheet = sheet;
        }

        private static void SetBorderStyle(ICellStyle style, BorderStyle topStyle, BorderStyle rightStyle,
            BorderStyle bottomStyle, BorderStyle leftStyle)
        {
            style.BorderTop = topStyle;
            style.BorderRight = rightStyle;
            style.BorderBottom = bottomStyle;
            style.BorderLeft = leftStyle;
        }

        /// <summary>
        /// get style by format string
        /// NPOI not allow styles with duplicated formatstring
        /// </summary>
        /// <param name="cellDataFormatString"></param>
        /// <returns></returns>
        private ICellStyle BuildCellDataFormatCache(string cellDataFormatString)
        {

            ICellStyle style;
            if (!_cellFormatCache.ContainsKey(cellDataFormatString))
            {
                style = _sheet.Workbook.CreateCellStyle();
                int formatId = HSSFDataFormat.GetBuiltinFormat(cellDataFormatString);
                // not a builtin data format
                if (formatId == -1)
                {
                    // create new format only when style with same format string not find in cache
                    var format = _sheet.Workbook.CreateDataFormat();
                    style.DataFormat = format.GetFormat(cellDataFormatString);
                }
                else
                {
                    style.DataFormat = (short) formatId;
                }
                _cellFormatCache[cellDataFormatString] = style;
            }
            else
            {
                style = _cellFormatCache[cellDataFormatString];
            }
            return style;
        }


        /// <summary>
        /// get column attribute from cache
        /// </summary>
        /// <param name="headerText"></param>
        /// <returns></returns>
        private ExileColumnDataFormatAttribute BuildExilePropertyAttribute(string headerText)
        {
            ExileColumnDataFormatAttribute attribute;
            if (!_dataFormatAttr.ContainsKey(headerText))
            {
                attribute = Utils.GetTypeAttribute(typeof (T), headerText) as ExileColumnDataFormatAttribute;
                _dataFormatAttr.Add(headerText, attribute);
            }
            else
            {
                attribute = _dataFormatAttr[headerText];
            }
            return attribute;
        }

        public void FillContent(IList<T> dataList)
        {
            var currentRowIndex = 0;

            // todo move to attribute
            // set print margin
            //_sheet.SetMargin(MarginType.TopMargin, 0);
//            _sheet.SetMargin(MarginType.BottomMargin, 0);
//            _sheet.SetMargin(MarginType.LeftMargin, 0);
//            _sheet.SetMargin(MarginType.RightMargin, 0);
            _sheet.PrintSetup.Scale = 100;
            // set paper size as A4
            _sheet.PrintSetup.PaperSize=9;

            // todo remove visibility from here use it from headerAttribute
            // bulid sheet title
            var titleRow = _sheet.CreateRow(currentRowIndex);
            currentRowIndex = 1;

            var titleCell = titleRow.CreateCell(0);
            var headerstyle = _sheet.Workbook.CreateCellStyle();
            headerstyle.WrapText = true;
            headerstyle.Alignment = HorizontalAlignment.Center;
            headerstyle.VerticalAlignment = VerticalAlignment.Center;

            var font = _sheet.Workbook.CreateFont();
            font.FontHeightInPoints = DocumentMeta.TitleFontHeight;
            headerstyle.SetFont(font);
            titleRow.HeightInPoints = DocumentMeta.TitleRowHeight;

            titleCell.CellStyle = headerstyle;
            titleCell.SetCellValue(DocumentMeta.TitleText);

            var cra = new CellRangeAddress(0, 0, 0, DocumentMeta.Headers.Count - 1);
            _sheet.AddMergedRegion(cra);


            // build sheet header
            var headerRow = _sheet.CreateRow(currentRowIndex);
            var columnIndex = 0;
            currentRowIndex++;
            foreach (var h in DocumentMeta.Headers)
            {
                var cell = headerRow.CreateCell(columnIndex);
                var style = _sheet.Workbook.CreateCellStyle();
                // set border
                SetBorderStyle(style, BorderStyle.Thin, BorderStyle.Thin, BorderStyle.Thin,
                    BorderStyle.Thin);

                cell.CellStyle = style;
                cell.SetCellValue(h.PropertyDescription);
                if (h.Width != 0)
                {
                    _sheet.SetColumnWidth(columnIndex, h.Width * 256);
                }
                columnIndex++;
            }


            // build sheet data
            for (int i = 0; i < dataList.Count; i++)
            {
                var dataRow = _sheet.CreateRow(i + currentRowIndex);
                //set row height
//                                if (DocumentMeta.RowHeight!=0)
//                                {
//                                    dataRow.HeightInPoints = DocumentMeta.RowHeight;
//                                }

                columnIndex = 0;
                foreach (var h in DocumentMeta.Headers)
                {
                    var cell = dataRow.CreateCell(columnIndex);
                    var value = Utils.GetPropValue(dataList[i], h.PropertyName);
                    var type = Utils.GetPropType(dataList[i], h.PropertyName);

                    if (h.ColumnType == ExileColumnType.AutoIndex)
                    {
                        cell.SetCellValue(i + 1);
                    }
                    else
                    {
                        // set cell value
                        if (type == typeof(DateTime))
                        {
                            cell.SetCellValue((DateTime)value);
                        }
                        else if (type == typeof(double) || type == typeof(Single) || type == typeof(float) ||
                                 type == typeof(int))
                        {
                            cell.SetCellValue(Convert.ToDouble(value));
                        }
                        else
                        {
                            cell.SetCellValue(value.ToString());
                        }
                    }

                    // set cell data format
                    // explicitly defined on ExilePropertyAttribute
                    string formatString;
                    var dataFormatAttr = BuildExilePropertyAttribute(h.PropertyName);
                    if (dataFormatAttr != null)
                    {
                        var formatId = (int)dataFormatAttr.ColumnBulitinDataFormat;
                        formatString = string.IsNullOrEmpty(dataFormatAttr.ColumnCustomDataFormat)
                            ? BuiltinFormats.GetBuiltinFormat(formatId)
                            : dataFormatAttr.ColumnCustomDataFormat;
                    }
                    else
                    {
                        formatString = ExcelDataFormat.Conver2ExcelDataFormat(type);
                    }

                    var style = BuildCellDataFormatCache(formatString);
                    // set auto wrap
                    if (h.AutoFit)
                    {
                        style.WrapText = true;
                    }

                    cell.CellStyle = style;

                    // set border
                    SetBorderStyle(cell.CellStyle, BorderStyle.Thin, BorderStyle.Thin, BorderStyle.Thin,
                        BorderStyle.Thin);

                    columnIndex++;
                }
            }
        }
    }


}

