using NPOI.SS.Formula.Functions;
using NPOI.SS.Util;
using NPOI.Util;

namespace ExileExcel
{
    using ExileExcel.Attribute;
    using ExileExcel.Common;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using System;
    using System.Collections.Generic;
    using System.IO;

    public class ExileExtractor<T>
    {
        private readonly Dictionary<string, ICellStyle> _cellFormatCache;
        private readonly Dictionary<string, ExilePropertyAttribute> _exilePropertyAttributes;

        private IWorkbook _workbook;
        private ISheet _sheet;

        private SheetAttributeOverride sheetAttribute;

        /// <summary>
        /// Init
        /// </summary>
        public ExileExtractor()
        {
            // fill all buildin dataformat for avoid re-declare same format
            _cellFormatCache = new Dictionary<string, ICellStyle>();
            _exilePropertyAttributes = new Dictionary<string, ExilePropertyAttribute>();
            sheetAttribute=new SheetAttributeOverride();
        }

        public ExileExtractor(string sheetName,string titleText):this()
        {
            sheetAttribute.SheetName = sheetName;
            sheetAttribute.TitleText = titleText;
        } 

        private struct SheetAttributeOverride
        {
            public string SheetName;
            public string TitleText;
        }

        public void ExcelWriteStream(List<T> dataList, Stream stream, ExileExtractTypes extractType)
        {
            // check type matched
            var matchedType = ParserUtilty.GetTypeMatched<T>();
            if (matchedType == null)
            {
                throw new ArgumentNullException("dataList");
            }

            // override sheetname title
            OverrideSheetAttribute(matchedType);
            _workbook = BuildWorkbook(dataList, extractType, matchedType);

            _workbook.Write(stream);
        }

        private void OverrideSheetAttribute(ExileMatchResult matched)
        {
           if(!string.IsNullOrEmpty(sheetAttribute.SheetName)) matched.SheetName = sheetAttribute.SheetName;
           if(!string.IsNullOrEmpty(sheetAttribute.TitleText)) matched.TitleText = sheetAttribute.TitleText;
        }

        private IWorkbook BuildWorkbook(IList<T> dataList, ExileExtractTypes extractType, ExileMatchResult matchedType)
        {
            switch (extractType)
            {
                case ExileExtractTypes.Excel2003:
                    _workbook = new HSSFWorkbook();
                    break;
                case ExileExtractTypes.Excel2007:
                    _workbook = new XSSFWorkbook();
                    break;
                default:
                    throw new ArgumentOutOfRangeException("extractType");
            }

            // build sheet attributes
            _sheet = _workbook.CreateSheet(matchedType.SheetName);
            var currentRowIndex = 0;

            // bulid sheet title
            var titleVisibility = matchedType.Visibility;
            if (titleVisibility!=ExileHeaderVisibility.Invisible)
            {
                var titleRow = _sheet.CreateRow(currentRowIndex);
                currentRowIndex = 1;
                    
                var titleCell = titleRow.CreateCell(0);
                var style = _workbook.CreateCellStyle();
                style.WrapText = true;
                style.Alignment = HorizontalAlignment.CENTER;
                style.VerticalAlignment=VerticalAlignment.CENTER;

                var font = _workbook.CreateFont();
                font.FontHeight = matchedType.FontHeight;
                style.SetFont(font);
                titleRow.Height = matchedType.RowHeight;
           
                titleCell.CellStyle = style;
                titleCell.SetCellValue(matchedType.TitleText);

                if (titleVisibility==ExileHeaderVisibility.VisibleWithCellCombine)
                {
                    var cra = new CellRangeAddress(0, 0, 0, matchedType.Headers.Count - 1);
                    _sheet.AddMergedRegion(cra);
                }
            }

            // build sheet header
            var headerRow = _sheet.CreateRow(currentRowIndex);
            var columnIndex = 0;
            currentRowIndex++;
            foreach (var h in matchedType.Headers)
            {
                var cell = headerRow.CreateCell(columnIndex);
                var style = _workbook.CreateCellStyle();
                // set border
                SetBorderStyle(style, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN,
                    BorderStyle.THIN);
                cell.CellStyle = style;
                cell.SetCellValue(h.PropertyDescription);
                columnIndex++;
            }

            // build sheet data
            for (int i = 0; i < dataList.Count; i++)
            {
                var dataRow = _sheet.CreateRow(i + currentRowIndex);
                columnIndex = 0;
                foreach (var h in matchedType.Headers)
                {
                    var cell = dataRow.CreateCell(columnIndex);
                    var value = ParserUtilty.GetPropValue(dataList[i], h.PropertyName);
                    var type = ParserUtilty.GetPropType(dataList[i], h.PropertyName);

                    if (h.ColumnType==ExileColumnType.AutoIndex)
                    {
                        cell.SetCellValue(i+1);
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
                    var formatId = (int) BuildExilePropertyAttribute(h.PropertyName).ColumnBulitinDataFormat;
                    var formatString = BuildExilePropertyAttribute(h.PropertyName).ColumnCustomDataFormat;


                    if (!string.IsNullOrEmpty(formatString))
                    {

                    }
                    else if (formatId != 0)
                    {
                        formatString = BuiltinFormats.GetBuiltinFormat(formatId);
                    }
                    else
                    {
                        formatString = NPOIDataFormat.Conver2ExcelDataFormat(type);
                    }


                    cell.CellStyle = BuildCellDataFormatCache(formatString);

                    // set border
                    SetBorderStyle(cell.CellStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN,
                        BorderStyle.THIN);

                    columnIndex++;
                }
            }

            return _workbook;
        }

        private void SetBorderStyle(ICellStyle style, BorderStyle topStyle, BorderStyle rightStyle,
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
                style = _workbook.CreateCellStyle();
                int formatId = HSSFDataFormat.GetBuiltinFormat(cellDataFormatString);
                // not a builtin data format
                if (formatId == -1)
                {
                    // create new format only when style with same format string not find in cache
                    var format = _workbook.CreateDataFormat();
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
        private ExilePropertyAttribute BuildExilePropertyAttribute(string headerText)
        {
            ExilePropertyAttribute attribute = null;
            if (!_exilePropertyAttributes.ContainsKey(headerText))
            {
                attribute = ParserUtilty.GetTypeAttribute(typeof (T), headerText);

                _exilePropertyAttributes.Add(headerText, attribute);

            }
            else
            {
                attribute = _exilePropertyAttributes[headerText];
            }

            return attribute;
        }
    }
}
