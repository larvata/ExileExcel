using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using NPOI.HSSF.UserModel;


namespace ExileExcel
{
    public class ExcelParser
    {
        public string Description { get; set; }

        public List<BaseExileData> ParseFromXLS(string filePath)
        {
            HSSFWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
            }
            var worksheet = (HSSFSheet) workbook.GetSheetAt(0);
            var rows = worksheet.GetRowEnumerator();

            var matchedRawDataObject = new BaseExileData();
            var inputKeyPair = new Dictionary<int, string>();



            // get header
            rows.MoveNext();
            var header = (HSSFRow) rows.Current;
            for (int i = 0; i < header.Cells.Count; i++)
            {
                inputKeyPair.Add(i, header.Cells[i].StringCellValue);
            }

            var list = ReflectiveEnumerator.GetEnumerableOfType<BaseExileData>();
            foreach (var rawDataObject in list.Where(r => r.IsKeyPairMatch(inputKeyPair)))
            {
                matchedRawDataObject = rawDataObject;
                Description = matchedRawDataObject.GetClassDescription();
                break;
            }

            var retList = new List<BaseExileData>();

            // get body
            while (rows.MoveNext())
            {
                var tmpRawData = (BaseExileData) Activator.CreateInstance(matchedRawDataObject.GetType());
                {
                    var row = (HSSFRow) rows.Current;
                    foreach (var p in tmpRawData.GetNameAttributePair())
                    {
                        var index = inputKeyPair.FirstOrDefault(k => k.Value == p.Value).Key;
                        var prop = tmpRawData.GetType()
                                             .GetProperty(p.Key,
                                                          BindingFlags.NonPublic | BindingFlags.Public |
                                                          BindingFlags.Instance);
                        var cell = (HSSFCell) row.GetCell(index);

                        if (null == prop || !prop.CanWrite || null == cell) continue;

                        if (prop.PropertyType == typeof (int))
                        {
                            prop.SetValue(tmpRawData, Convert.ToInt32(cell.ToString()), null);
                        }
                        else
                        {
                            prop.SetValue(tmpRawData, cell.ToString(), null);
                        }
                    }
                    retList.Add(tmpRawData);
                }

            }


            return retList;
        }
    }

}