namespace ExileExcel.Attribute
{
    using System;
    using System.ComponentModel;
    using ExileExcel.Common;

    [Description("RawData Base Class")]
    [AttributeUsage(AttributeTargets.Property)]
    public class ExilePropertyAttribute : System.Attribute
    {
        private readonly string _headerText;
        private readonly ExileColumnType _columnType;
        private readonly int _headerTextSequence;
        private readonly NPOIDataFormatEnum _columnBulitinDataFormat;
        private readonly string _columnCustomDataFormat;


        /// <summary>
        /// Excel sheet header text
        /// </summary>
        public string HeaderText
        {
            get { return _headerText; }
        }

        /// <summary>
        /// if true, output index of entity instead of property value
        /// </summary>
        public ExileColumnType ColumnType
        {
            get { return _columnType; }
        }

        /// <summary>
        /// Column data format, should be excel builtin format
        /// </summary>
        public NPOIDataFormatEnum ColumnBulitinDataFormat
        {
            get { return _columnBulitinDataFormat; }
        }

        /// <summary>
        /// Custom column data format like '0.00' for 2 decimal places number
        /// </summary>
        public string ColumnCustomDataFormat
        {
            get { return _columnCustomDataFormat; }
        }

        /// <summary>
        /// Output sequence
        /// RULE: HeaderTextSequence!=-1 ASC + HeaderTextSequence==-1 
        /// </summary>
        public int HeaderTextSequence
        {
            get { return _headerTextSequence; }
        }

        public ExilePropertyAttribute()
        {
           _headerText=string.Empty;
           _columnType=ExileColumnType.CommonData;
           _headerTextSequence=-1;
           _columnBulitinDataFormat=NPOIDataFormatEnum.Text;
           _columnCustomDataFormat=string.Empty;
        }

        public ExilePropertyAttribute(
            string headerText,
            ExileColumnType columnType,
            NPOIDataFormatEnum columnBulitinDataFormat = NPOIDataFormatEnum.Text,
            string columnCustomDataFormat = "",
            int headerTextSequence=-1 
            ):this()
        {
            _headerTextSequence = headerTextSequence;
            _headerText = headerText;
            _columnType = columnType;
            _columnBulitinDataFormat = columnBulitinDataFormat;
            _columnCustomDataFormat = columnCustomDataFormat;

        }

        public ExilePropertyAttribute(
            string headerText, 
            NPOIDataFormatEnum columnBulitinDataFormat = NPOIDataFormatEnum.Text
            ):this()
        {
            _headerText = headerText;
            _columnBulitinDataFormat = columnBulitinDataFormat;
        }

        public ExilePropertyAttribute(
            string headerText,
            string columnCustomDataFormat = ""
            ):this()
        {
            _headerText = headerText;
            _columnCustomDataFormat = columnCustomDataFormat;
        }

        public ExilePropertyAttribute(
            string headerText,
            ExileColumnType columnType=ExileColumnType.CommonData
            ):this()
        {
            _headerText = headerText;
            _columnType = columnType;
        }

        public ExilePropertyAttribute(
            string headerText
            )
            : this()
        {
            _headerText = headerText;
        }
    }
}
