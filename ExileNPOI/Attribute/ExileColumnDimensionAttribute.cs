namespace ExileNPOI.Attribute
{
    using System;
    using System.ComponentModel;

    [Description("Column size and presention")]
    [AttributeUsage(AttributeTargets.Property)]
    public class ExileColumnDimensionAttribute:System.Attribute
    {
        public int ColumnWidth { get; set; }
        public int RowHeight { get; set; }
        public bool AutoFit { get; set; }
    }
}
