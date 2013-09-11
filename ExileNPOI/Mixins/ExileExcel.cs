using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace ExileNPOI.Mixins
{
    public static class MojitoCells
    {
//        public static void Parse<T>(this Workbook workbook, string filePath) where T : IMojitoExcel   
//        {
//            
//        }


        public static void Compose<T>(this ISheet sheet, IList<T> list) where T : IExilable
        {
            var extractor = new ExileExcelExtractor<T>(sheet);
            extractor.FillContent(list);
        }

        public static void Compose<T>(this ISheet sheet, IList<T> list, string titleText) where T : IExilable
        {
            var extractor = new ExileExcelExtractor<T>(sheet);
            extractor.DocumentMeta.TitleText = titleText;
            extractor.FillContent(list);
        }

//        /// <summary>
//        /// 
//        /// </summary>
//        /// <typeparam name="T"></typeparam>
//        /// <param name="workbook"></param>
//        /// <param name="list"></param>
//        /// <returns>worksheet id</returns>
//        public static Worksheet Compose<T>(this Workbook workbook, IList<T> list) where T : IMojitoExcel
//        {
//            // memo
//            // sheet name defined by MojitoSheetAttribute
//
//            var sheetId=workbook.Worksheets.Add();
//
//            return Compose(workbook, list, sheetId);
//
//
//        }

        
//
//        public static void Compose<T>(this Aspose.Cells.Workbook workbook, IList<T> list,string sheetName) where T : IMojitoExcel
//        {
//            
//        }

    }
}
