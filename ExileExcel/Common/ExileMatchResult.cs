namespace ExileExcel.Common
{
    using System;
    using System.Collections.Generic;

    public class ExileMatchResult
    {

        public List<ExileHeader> Headers;
        public Type MatchedType { get; set; }
        public string SheetName { get; set; }
        public string TitleText { get; set; }
        public ExileHeaderVisibility Visibility { get; set; }
        public short FontHeight { get; set; }
        public short RowHeight { get; set; }

        public ExileMatchResult()
        {
            Headers = new List<ExileHeader>();
        }
    }

    public struct ExileHeader
    {
        public string PropertyName;
        public string PropertyDescription;
        public int DisplaySequence;
        public NPOIDataFormatEnum BuiltinFormat;
        public string CustomDataFormat;
        public ExileColumnType ColumnType { get; set; }

    }
}
