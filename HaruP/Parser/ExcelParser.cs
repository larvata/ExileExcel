using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using HaruP.Common;
using NPOI.SS.UserModel;

namespace HaruP.Parser
{
    public class ExcelParser
    {
        
        public string Description { get; set; }

        public List<IBaseSheetData> Parse(string filePath)
        {
            IWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }
            var worksheet = workbook.GetSheetAt(0);
            var rows = worksheet.GetRowEnumerator();

            IBaseSheetData matchedRawDataObject = null;
            var inputKeyPair = new Dictionary<int, string>();



            // get header
            rows.MoveNext();
            var header = (IRow)rows.Current;
            for (int i = 0; i < header.Cells.Count; i++)
            {
                inputKeyPair.Add(i, header.Cells[i].StringCellValue);
            }

            var list = Utils.GetEnumerableOfType<IBaseSheetData>();
            foreach (var rawDataObject in list.Where(r => r.IsKeyPairMatch(inputKeyPair)))
            {
                matchedRawDataObject = rawDataObject;
                Description = matchedRawDataObject.GetClassDescription();
                break;
            }

            var retList = new List<IBaseSheetData>();

            // get body
            while (rows.MoveNext())
            {
                var tmpRawData = (IBaseSheetData)Activator.CreateInstance(matchedRawDataObject.GetType());
                {
                    var row = (IRow)rows.Current;
                    foreach (var p in tmpRawData.GetNameAttributePair())
                    {
                        var index = inputKeyPair.FirstOrDefault(k => k.Value == p.Value).Key;
                        var prop = tmpRawData.GetType()
                                             .GetProperty(p.Key,
                                                          BindingFlags.NonPublic | BindingFlags.Public |
                                                          BindingFlags.Instance);
                        var cell = row.GetCell(index);

                        if (null == prop || !prop.CanWrite || null == cell) continue;

                        if (prop.PropertyType == typeof(int))
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
