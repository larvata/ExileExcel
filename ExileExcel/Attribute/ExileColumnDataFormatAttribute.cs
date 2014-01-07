using System;
using System.ComponentModel;
using ExileExcel.Common;

namespace ExileExcel.Attribute

{
    /// <summary>
    /// Mark class can be find by ExileParser
    /// </summary>
    [Description("Header Column Style Define")]
    [AttributeUsage(AttributeTargets.Property)]
    public class ExileColumnDataFormatAttribute : System.Attribute
    {

        /// <summary>
        /// if true, output index of entity instead of property value
        /// </summary>
        public ExileColumnType ColumnType { get; set; }


        /// <summary>
        /// Column data format, should be excel builtin format
        /// </summary>
        public ExcelDataFormatEnum ColumnBulitinDataFormat { get; set; }


        /// <summary>
        /// Custom column data format like '0.00' for 2 decimal places number
        /// </summary>
        public string ColumnCustomDataFormat { get; set; }



    }
}
