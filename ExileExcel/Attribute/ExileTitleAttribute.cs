using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace ExileExcel.Attribute
{
    /// <summary>
    /// Mark class can be find by ExileParser
    /// </summary>
    [Description("Exiliable Title Attribute")]
    [AttributeUsage(AttributeTargets.Class)]
    public class ExileTitleAttribute : System.Attribute
    {
        private readonly string _titleText;
        private readonly ExileHeaderVisibility _visibility;


        public string TitleText
        {
            get
            {
                return _titleText;
            }
        }

        public bool IsCombineTitleCell
        {
            get
            {
                return _isCombineTitleCell;
            }
        }

        public bool IsDisplayTitle
        {
            get
            {
                return _isDisplayTitle;
            }
        }

        public int FontHeight
        {
            get { return _fontHeight; }
        }

        public short RowHeight
        {
            get { return _rowHeight; }
        }

        public ExileTitleAttribute(string titleText = "", bool isCombineTitleCell = true, bool isDisplayTitle = false,
            int fontHeight = -1, short rowHeight = -1)
        {
            _titleText = titleText;
            _isCombineTitleCell = isCombineTitleCell;
            _isDisplayTitle = isDisplayTitle;
            _fontHeight = fontHeight;
            _rowHeight = rowHeight;

        }
    }
}
