using System;
using System.Collections.Generic;
using ExileExcel.Attribute;

namespace ExileExcel.Common
{
    /// <summary>
    /// exile document Modal
    /// </summary>
    public class ExileDocumentMeta
    {
        internal Type MatchedType { get; set; }
        public List<ExileHeaderMeta> Headers;
        public string SheetName { get; set; }
        public string TitleText { get; set; }
        public short TitleFontHeight { get; set; }
        public short TitleRowHeight { get; set; }

        public int StartRowNum { get; set; }
        public int StartColumnNum { get; set; }

        public bool HideHeader { get; set; }

        public ExileDocumentMeta()
        {
            Headers = new List<ExileHeaderMeta>();
        }
    }

    public struct ExileHeaderMeta
    {
        public string PropertyName;
        public string PropertyDescription;
        public int DisplaySequence;
        public ExcelDataFormatEnum BuiltinFormat;
        public string CustomDataFormat;
        public ExileColumnType ColumnType;
        public int Width;
        public int Height;
        public bool AutoFit;
    }
}
