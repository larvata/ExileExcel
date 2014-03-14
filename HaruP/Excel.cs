using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HaruP.Common;
using NPOI.SS.UserModel;

namespace HaruP
{
    public abstract class Excel
    {
        // npoi
        internal IWorkbook workbook;
        internal ISheet sheet;
        internal SheetMeta sheetMeta;
    }
}
