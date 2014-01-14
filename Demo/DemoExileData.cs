using System;
using ExileExcel;
using ExileExcel.Attribute;
using NPOI.SS.Formula.Functions;

namespace Demo
{
    [ExileSheetTitle(RowHeight = 40, FontHeight = 40)]
    [ExileSheetDataArea(StartRowNum = 5)]
    class DemoExileData:IExilable
    {
        [ExileColumnDataFormat(ColumnType = ExileColumnType.AutoIndex)]
        [ExileHeaderGeneral(HeaderText = "序号", HeaderSequence = 1)]
        public int Id { get; set; }

        [ExileColumnDimension(AutoFit = true,ColumnWidth = 20)]
        [ExileHeaderGeneral(HeaderText = "学号", HeaderSequence = 1)]
        public string Number { get; set; }

        [ExileHeaderGeneral(HeaderText = "姓名", HeaderSequence = 2)]
        public string Name { get; set; }

        [ExileHeaderGeneral(HeaderText = "分数", HeaderSequence = 3)]
        public float Score { get; set; }

        [ExileColumnDataFormat(ColumnCustomDataFormat = "YYYY/MM/DD HH:mm:ss")]
        [ExileHeaderGeneral(HeaderText = "考试日期", HeaderSequence = 4)]
        [ExileColumnDimension(AutoFit = true,ColumnWidth = 20)]
        public DateTime TestDate { get; set; }
    }

}
