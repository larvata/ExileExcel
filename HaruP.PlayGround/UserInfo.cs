using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HaruP.Parser;

namespace HaruP.PlayGround
{
    public class UserInfo : IBaseSheetData
    {
        [SheetData("Id",true)]
        public int Id { get; set; }
        [SheetData("Name", true)]
        public string Name { get; set; }
        [SheetData("Age", true)]
        public int Age { get; set; }
        [SheetData("Mail", true)]
        public string Email { get; set; }
    }
}
