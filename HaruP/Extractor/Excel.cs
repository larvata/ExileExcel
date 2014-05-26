using HaruP.Common;
using NPOI.SS.UserModel;

namespace HaruP.Extractor
{
    public abstract class Excel
    {
        // npoi
        internal IWorkbook workbook;
        internal ISheet sheet;
        internal SheetMeta sheetMeta;
    }
}
