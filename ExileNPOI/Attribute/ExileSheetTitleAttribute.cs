namespace ExileNPOI.Attribute
{
    using System;
    using System.ComponentModel;

    [Description("Header Cell Style Define")]
    [AttributeUsage(AttributeTargets.Class)]
    public class ExileSheetTitleAttribute:Attribute
    {
        public short FontHeight { get; set; }
        public short RowHeight { get; set; }
    }
}
