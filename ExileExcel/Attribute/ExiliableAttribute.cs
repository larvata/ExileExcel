using System.Reflection.Emit;
using ExileExcel.Common;

namespace ExileExcel.Attribute
{
    using System;
    using System.ComponentModel;

    /// <summary>
    /// Mark class can be find by ExileParser
    /// </summary>
    [Description("Exiliable Class Attribute")]
    [AttributeUsage(AttributeTargets.Class)]
    public class ExiliableAttribute : Attribute
    {
        // text of sheet
        private readonly string _sheetName;
        // indicate visibility of header(sheet header)  
        private readonly ExileHeaderVisibility _visibility;

        private readonly string _titleText;
        private readonly short _fontHeight;
        private readonly short _rowHeight;

        public string SheetName
        {
            get { return _sheetName; }
        }

        public ExileHeaderVisibility Visibility
        {
            get { return _visibility; }
        }

        public string TitleText
        {
            get { return _titleText; }
        }

        public short FontHeight
        {
            get { return _fontHeight; }
        }

        public short RowHeight
        {
            get { return _rowHeight; }
        }

        public ExiliableAttribute()
        {
            _sheetName = "Default Sheet";
            _titleText = "ExileExcel Untitled";
            _visibility = ExileHeaderVisibility.Invisible;
            _fontHeight = 20;
            _rowHeight = 400;
        }

        public ExiliableAttribute(string sheetName = "")
        {
            _sheetName = sheetName;
        }

        public ExiliableAttribute(ExileHeaderVisibility visiblity, string sheetName = "Default Sheet")
            : this()
        {
            _sheetName = sheetName;
            _visibility = visiblity;
        }

        public ExiliableAttribute(ExileHeaderVisibility visiblity, string sheetName,
            string titleText = "ExileExcel Untitled")
            : this()
        {
            _sheetName = sheetName;
            _visibility = visiblity;
            _titleText = titleText;
        }

        public ExiliableAttribute(ExileHeaderVisibility visiblity, string sheetName, string titleText,
            short fontHeight = 20, short rowHeight = 400)
        {
            _sheetName = sheetName;
            _visibility = visiblity;
            _titleText = titleText;
            _fontHeight = fontHeight;
            _rowHeight = rowHeight;
        }
    }
}
