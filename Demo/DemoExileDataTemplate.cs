using System;
using ExileExcel;
using ExileExcel.Attribute;
using NPOI.SS.Formula.Functions;

namespace Demo
{
    [ExileSheetDataArea(StartRowNum = 3)]
    [ExileSheetTitle(HideHeader=true )]
    public class DemoExileDataTemplate : IExilable
    {

        [ExileColumnDataFormat(ColumnType = ExileColumnType.AutoIndex)]
        [ExileHeaderGeneral(HeaderSequence = 1)]
        public int Id { get; set; }

        [ExileColumnDimension(AutoFit = true, ColumnWidth = 20)]
        [ExileHeaderGeneral(HeaderSequence = 1)]
        public string Number { get; set; }

        [ExileHeaderGeneral(HeaderSequence = 2)]
        public string Name { get; set; }

        [ExileHeaderGeneral(HeaderSequence = 3)]
        public float Score { get; set; }

        [ExileColumnDataFormat(ColumnCustomDataFormat = "YYYY/MM/DD HH:mm:ss")]
        [ExileHeaderGeneral(HeaderText = "考试日期", HeaderSequence = 4)]
//        [ExileColumnDimension(AutoFit = true, ColumnWidth = 20)]
       public DateTime TestDate { get; set; }
    }

}
