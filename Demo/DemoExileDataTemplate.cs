using System;
using ExileExcel;
using ExileExcel.Attribute;

namespace Demo
{
    [ExileSheetDataArea(StartRowNum = 2)]
    [ExileSheetTitle(HideHeader=true )]
    public class DemoExileDataTemplate : IExilable
    {

        [ExileColumnDataFormat(ColumnType = ExileColumnType.AutoIndex)]
        [ExileHeaderGeneral(HeaderSequence = 1)]
        public int Id { get; set; }

        [ExileColumnDimension(AutoFit = true, ColumnWidth = 20)]
        [ExileHeaderGeneral(HeaderSequence = 2)]
        public string Number { get; set; }

        [ExileHeaderGeneral(HeaderSequence = 3)]
        public string Name { get; set; }

        [ExileHeaderGeneral(HeaderSequence = 4)]
        public float Score { get; set; }

        [ExileColumnDataFormat(ColumnCustomDataFormat = "YYYY/MM/DD HH:mm:ss")]
        [ExileHeaderGeneral(HeaderSequence = 5)]
        public DateTime TestDate { get; set; }
    }

}
