using System;
using System.ComponentModel;

namespace ExileExcel.Attribute
{
    [Description("Header Cell Style Define")]
    [AttributeUsage(AttributeTargets.Class)]
    public class ExileSheetTitleAttribute:System.Attribute
    {
        public short FontHeight { get; set; }
        public short RowHeight { get; set; }
    }
}
