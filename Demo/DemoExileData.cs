using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExileExcel;
using ExileExcel.Attribute;

namespace Demo
{
    [ExiliableAttribute("Demo class")]
    class DemoExileData
    {
        public int Id { get; set; }
        [ExileProperty("学号")]
        public string Number { get; set; }
        [ExileProperty("姓名")]
        public string Name { get; set; }
        [ExileProperty("分数")]
        public float Score { get; set; }
        [ExileProperty("考试日期")]
        public DateTime TestDate { get; set; }
    }

}
