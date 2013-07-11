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
        private readonly bool _identityColumn;
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
        /// may never used
        /// </summary>
        public bool IdentityColumn
        {
            get { return _identityColumn; }
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

        public ExilePropertyAttribute(
            string headerText,
            string columnCustomDataFormat="",
            NPOIDataFormatEnum columnBulitinDataFormat = NPOIDataFormatEnum.Text, 
            int headerTextSequence=-1, 
            bool identityColumn = false)
        {
            _headerTextSequence = headerTextSequence;
            _headerText = headerText;
            _identityColumn = identityColumn;
            _columnBulitinDataFormat = columnBulitinDataFormat;
            _columnCustomDataFormat = columnCustomDataFormat;

        }
    }
}
