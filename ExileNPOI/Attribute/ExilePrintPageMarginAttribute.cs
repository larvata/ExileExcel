namespace ExileNPOI.Attribute
{
    using System;
    using System.ComponentModel;

    [Description("Page margin setting")]
    [AttributeUsage(AttributeTargets.Property)]
    class ExilePrintPageMarginAttribute:Attribute
    {
        public int MarginLeft;
        public int MarginTop;
        public int MarginRight;
        public int MarginBottom;
    }
}
