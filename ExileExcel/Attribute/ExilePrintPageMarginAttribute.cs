using System;
using System.ComponentModel;

namespace ExileExcel.Attribute
{
    [Description("Page margin setting")]
    [AttributeUsage(AttributeTargets.Property)]
    class ExilePrintPageMarginAttribute:System.Attribute
    {
        public int MarginLeft;
        public int MarginTop;
        public int MarginRight;
        public int MarginBottom;
    }
}
