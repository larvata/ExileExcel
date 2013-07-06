using System.Data;
using NPOI.SS.Formula.Functions;

namespace ExileExcel
{
    using ExileExcel.Common;
    using NPOI.SS.UserModel;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    public class ExcelParser<T>
    {

        /// <summary>
        /// Schema of matched class
        /// </summary>
        public ExileMatchResult MatchedSchema { get; set; }

        #region Private parser method

        private List<T> ParseFromExcel(string filePath,int sheetCount,string sheetName)
        {
            // todo verify filepath is exist 

            // open file Excel
            IWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }

            // read data from sheet by sheetCount OR sheetName
            ISheet worksheet;
            if (sheetCount!=-1)
            {
                worksheet = workbook.GetSheetAt(sheetCount);
            }
            else if (sheetName.Length>0)
            {
                worksheet = workbook.GetSheet(sheetName);
            }
            else
            {
                worksheet = workbook.GetSheetAt(0);
            }
            var rows = worksheet.GetRowEnumerator();

            // get sheet header text
            var inputKeyPair = new Dictionary<int, string>();
            rows.MoveNext();
            var header = (IRow) rows.Current;
            for (int i = 0; i < header.Cells.Count; i++)
            {
                inputKeyPair.Add(i, header.Cells[i].StringCellValue);
            }

            // check all ExileDataAttribute in T is matched with sheet header text
            MatchedSchema = ParserUtilty.GetTypeMatched<T>(inputKeyPair);
            if (MatchedSchema == null)
            {
                throw new Exception("RESULT NOT FOUND");
            }

            var retList = new List<T>();

            // aquire data from sheet
            while (rows.MoveNext())
            {
                var tmpRawData = (T)Activator.CreateInstance(typeof(T));
                var row = (IRow) rows.Current;
                foreach (var p in MatchedSchema.HeaderKeyPair)
                {
                    var index = inputKeyPair.FirstOrDefault(k => k.Value == p.Value).Key;
                    var prop = tmpRawData.GetType()
                                    .GetProperty(p.Key, BindingFlags.Public |BindingFlags.Instance);
                    var cell = row.GetCell(index);

                    if (null == prop || !prop.CanWrite || null == cell) continue;

                    // assert property with different type
                    if (prop.PropertyType == typeof(int))
                    {
                        prop.SetValue(tmpRawData, Convert.ToInt32(cell.NumericCellValue), null);
                    }
                    else if (prop.PropertyType==typeof(Single))
                    {
                        prop.SetValue(tmpRawData, Convert.ToSingle(cell.ToString()), null);
                    }
                    else if (prop.PropertyType==typeof(DateTime))
                    {
                        prop.SetValue(tmpRawData,cell.DateCellValue,null);
                    }
                    else 
                    {
                        prop.SetValue(tmpRawData, cell.ToString(), null);
                    }
                }
                retList.Add(tmpRawData);
            }

            return retList;
        }

        private List<T> ParseFromDataSet(DataSet inputDataSet)
        {
            throw new NotImplementedException();
        } 
        #endregion

        #region Public parser method

        /// <summary>
        /// Parse from DataSet
        /// </summary>
        /// <param name="inputDataSet">DataSet source</param>
        /// <returns></returns>
        public List<T> Parse(DataSet inputDataSet)
        {
            return ParseFromDataSet(inputDataSet);
        }

        /// <summary>
        /// Parse from excel file
        /// </summary>
        /// <param name="filePath">path of excel file</param>
        /// <returns>list of parsed data </returns>
        public List<T> Parse(string filePath)
        {
            return ParseFromExcel(filePath, 0,string.Empty);
        }

        /// <summary>
        /// Parse from excel file
        /// </summary>
        /// <param name="filePath">path of excel file</param>
        /// <param name="sheetCount">index of sheet (zero base)</param>
        /// <returns></returns>
        public List<T> Parse(string filePath, int sheetCount)
        {
            return ParseFromExcel(filePath, sheetCount, string.Empty);
        }

        /// <summary>
        /// Parse from excel file
        /// </summary>
        /// <param name="filePath">path of excel file</param>
        /// <param name="sheetName">name of sheet</param>
        /// <returns></returns>
        public List<T> Parse(string filePath, string sheetName)
        {
            return ParseFromExcel(filePath, -1, sheetName);
        }

        #endregion

        public MemoryStream ExcelOutputStream(List<T> dataList,Stream stream)
        {
            throw new NotImplementedException();
            //var bytes = new byte[stream.Length];
            //ms.Read(bytes, 0, (int)ms.Length);
            //fs.Write(bytes, 0, bytes.Length);
            //return null;


        }

    }
}