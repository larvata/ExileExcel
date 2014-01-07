using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExileExcel
{
    public class ExileExtractor<T> where T:IExilable
    {
        // npoi
        private IWorkbook _workbook;
        private ISheet _sheet;

        public string TitleText { get; set; }
        public string SheetName { get; set; }

        /// <summary>
        /// Init
        /// </summary>
        public ExileExtractor()
        {

        }

        public ExileExtractor(string sheetName,string titleText):this()
        {
            SheetName = sheetName;
            TitleText = titleText;
        }

        public void WriteStream(List<T> dataList, Stream stream, ExileExtractTypes extractType)
        {
            switch (extractType)
            {
                case ExileExtractTypes.Excel2003NPOI:
                case ExileExtractTypes.Excel2007NPOI:
                    ExcelWriteStream(dataList, stream, extractType);
                    break;
                case ExileExtractTypes.Excel2007OpenXML:
                    break;
                case ExileExtractTypes.Word2007OpenXML:
                    WordWriteStream(dataList, stream);
                    break;
                default:
                    throw new ArgumentOutOfRangeException("extractType");
            }
        }

        private void ExcelWriteStream(IList<T> dataList, Stream stream, ExileExtractTypes extractType)
        {

            switch (extractType)
            {
                case ExileExtractTypes.Excel2003NPOI:
                    _workbook = new HSSFWorkbook();
                    break;
                case ExileExtractTypes.Excel2007NPOI:
                    _workbook = new XSSFWorkbook();
                    break;
                case ExileExtractTypes.Excel2007OpenXML:
                    break;
                default:
                    throw new ArgumentOutOfRangeException("extractType");
            }

            _sheet = string.IsNullOrEmpty(SheetName) 
                ? _workbook.CreateSheet() 
                : _workbook.CreateSheet(SheetName);   

            var extractor = new ExileExcelExtractor<T>(_sheet);
            if (!string.IsNullOrEmpty(TitleText))
                extractor.DocumentMeta.TitleText = TitleText;
            
            extractor.FillContent(dataList);
       
            _workbook.Write(stream);
        }

        private void WordWriteStream(IList<T> dataList, Stream stream)
        {
            using(var ms=new MemoryStream())
            {
                using (var wordDocument = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    var body = new Body();
                    // ReSharper disable once PossiblyMistakenUseOfParamsMethod
                    mainPart.Document.Append(body);

                    var extractor = new ExileWordExtractor<T>(body);
                    if (!string.IsNullOrEmpty(TitleText))
                        extractor.DocumentMeta.TitleText = TitleText;
                    extractor.FillContent(dataList);

                    wordDocument.MainDocumentPart.Document.Save();
                }
                ms.Position = 0;
                ms.CopyTo(stream);
            }
        }

    }
}
