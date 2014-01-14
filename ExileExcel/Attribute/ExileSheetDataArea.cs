using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace ExileExcel.Attribute
{
    [Description("Data fill area")]
    [AttributeUsage(AttributeTargets.Class)]
    public class ExileSheetDataArea:System.Attribute
    {
        public int StartRowNum { get; set; }
        public int StartColumnNum { get; set; }
    }
}
