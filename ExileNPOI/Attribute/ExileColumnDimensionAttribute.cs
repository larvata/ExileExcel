using System;
using System.ComponentModel;

namespace ExileNPOI.Attribute
{
    [Description("Heae")]
    [AttributeUsage(AttributeTargets.Property)]
    public class ExileColumnDimensionAttribute:System.Attribute
    {
        public int ColumnWidth { get; set; }
        public int RowHeight { get; set; }
        public bool AutoFit { get; set; }
    }
}
