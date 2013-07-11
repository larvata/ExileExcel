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
        private readonly Dictionary<string,ICellStyle> _cellFormatCache;
        private readonly Dictionary<string, ExilePropertyAttribute> _exilePropertyAttributes;

        private IWorkbook _workbook ;
        private ISheet _sheet;

        /// <summary>
        /// Init
        /// </summary>
        public ExileExtractor()
        {
            // fill all buildin dataformat for avoid re-declare same format
            _cellFormatCache=new Dictionary<string, ICellStyle>();
            _exilePropertyAttributes=new Dictionary<string, ExilePropertyAttribute>();
        }

        public void ExcelWriteStream(List<T> dataList, Stream stream, ExtractType extractType)
        {
            // check type matched
            var matchedType = ParserUtilty.GetTypeMatched<T>();
            if (matchedType == null)
            {
                throw new ArgumentNullException("dataList");
            }
            _workbook = BuildWorkbook(dataList, extractType, matchedType);

            _workbook.Write(stream);
        }

        private IWorkbook BuildWorkbook(IList<T> dataList, ExtractType extractType, ExileMatchResult matchedType)
        {
            switch (extractType)
            {
                case ExtractType.Excel2003:
                    _workbook = new HSSFWorkbook();
                    break;
                case ExtractType.Excel2007:
                    _workbook = new XSSFWorkbook();
                    break;
                default:
                    throw new ArgumentOutOfRangeException("extractType");
            }

            _sheet = _workbook.CreateSheet();
            var headerRow = _sheet.CreateRow(0);

            // build sheet header
            var columnIndex = 0;
            foreach (var h in matchedType.HeaderKeyPair)
            {
                headerRow.CreateCell(columnIndex).SetCellValue(h.Value);
                columnIndex++;
            }

            // build sheet data
            for (int i = 0; i < dataList.Count; i++)
            {
                var dataRow = _sheet.CreateRow(i + 1);
                columnIndex = 0;
                foreach (var h in matchedType.HeaderKeyPair)
                {
                    var cell = dataRow.CreateCell(columnIndex);
                    var value = ParserUtilty.GetPropValue(dataList[i], h.Key);
                    var type = ParserUtilty.GetPropType(dataList[i], h.Key);

                    // set cell value
                    if (type == typeof (DateTime))
                    {
                        cell.SetCellValue((DateTime) value);
                    }
                    else if (type == typeof (double) || type == typeof (Single) || type == typeof (float) ||
                             type == typeof (int))
                    {
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else
                    {
                        cell.SetCellValue(value.ToString());
                    }


                    // set cell data format
                    // explicitly defined on ExilePropertyAttribute
                    var formatId = (int)BuildExilePropertyAttribute(h.Key).ColumnBulitinDataFormat;
                    var formatString = BuildExilePropertyAttribute(h.Key).ColumnCustomDataFormat;
                        

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


                    cell.CellStyle = BuildCellFormatCache(formatString);

                    columnIndex++;
                }
            }

            return _workbook;
        }

        /// <summary>
        /// get style by format string
        /// NPOI not allow styles with duplicated formatstring
        /// </summary>
        /// <param name="cellFormatString"></param>
        /// <returns></returns>
        private ICellStyle BuildCellFormatCache(string cellFormatString)
        {
            
            ICellStyle style;
            if (!_cellFormatCache.ContainsKey(cellFormatString))
            {
                style = _workbook.CreateCellStyle();
                int formatId = HSSFDataFormat.GetBuiltinFormat(cellFormatString);
                // not a buildin data format
                if (formatId == -1)
                {
                    // create new format only when style with same format string not find in cache
                    var format = _workbook.CreateDataFormat();
                    style.DataFormat = format.GetFormat(cellFormatString);
                }
                else
                {
                    style.DataFormat = (short)formatId;
                }
                _cellFormatCache[cellFormatString] = style;
            }
            else
            {
                style = _cellFormatCache[cellFormatString];
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
            ExilePropertyAttribute attribute=null;
            if (!_exilePropertyAttributes.ContainsKey(headerText))
            {
                 attribute = ParserUtilty.GetTypeAttribute(typeof(T),headerText);

                _exilePropertyAttributes.Add(headerText,attribute);

            }
            else
            {
                attribute = _exilePropertyAttributes[headerText];
            }

            return attribute;
        }
    }
}
