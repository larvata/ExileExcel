using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExileExcel.Common;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using ExileExcel.Common;

namespace ExileExcel
{
    public class ExcelExtractor<T>
    {

        public void ExcelWriteStream(List<T> dataList, Stream stream,ExtractType extractType)
        {
            // check type matched
            var matchedType = ParserUtilty.GetTypeMatched<T>();
            if (matchedType==null)
            {
                throw new ArgumentNullException("dataList");
            }
            var workbook = BuildWorkbook(dataList, extractType, matchedType);

            workbook.Write(stream);
        }

        private IWorkbook BuildWorkbook(List<T> dataList, ExtractType extractType,ExileMatchResult matchedType)
        {
            IWorkbook workbook = null;
            switch (extractType)
            {
                case ExtractType.Excel2003:
                    workbook=new HSSFWorkbook();
                    var sheet =new HSSFSheet((HSSFWorkbook)workbook);
                    //var row = sheet.CreateRow(0);
                    var headerRow = sheet.CreateRow(0);


                    
                    // build sheet header
                    var columnIndex = 0;
                    foreach (var h in matchedType.HeaderKeyPair)
                    {
                         headerRow.CreateCell(columnIndex).SetCellValue(h.Value);
                        columnIndex++;
                    }

                    // build sheet data
                    columnIndex = 0;
                    for (int i = 1; i < dataList.Count+1; i++)
                    {
                        var row = sheet.CreateRow(i);

                        foreach (var h in matchedType.HeaderKeyPair)
                        {
                            var cell=row.CreateCell(columnIndex);
                            cell.SetCellValue(ParserUtilty.GetPropValue(dataList[i],h.Key).ToString()); 
                        }
                        columnIndex++;
                    }

                    break;
                case ExtractType.Excel2007:
                    break;
                default:
                    throw new ArgumentOutOfRangeException("extractType");
            }
            return workbook;
        }
    }
}
