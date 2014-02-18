using System;
using System.Collections;
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
    public class ExileExtractor<T> where T : IExilable
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

        public ExileExtractor(string sheetName, string titleText)
            : this()
        {
            SheetName = sheetName;
            TitleText = titleText;
        }

        public void WriteStream(dynamic dataList, Stream stream, ExileExtractTypes extractType, string templateFilePath = "")
        {
            switch (extractType)
            {
                case ExileExtractTypes.Excel2003NPOI:
                case ExileExtractTypes.Excel2007NPOI:
                    ExcelWriteStream(dataList, stream, extractType, templateFilePath);
                    break;
                case ExileExtractTypes.Excel2007OpenXML:
                    break;
                case ExileExtractTypes.Word2007OpenXML:
                    WordWriteStream(dataList, stream, templateFilePath);
                    break;
                default:
                    throw new ArgumentOutOfRangeException("extractType");
            }

        }

        // Larvata  2014-02-18 15:59:10
        private void ExcelWriteStream(dynamic dataList, Stream stream, ExileExtractTypes extractType, string templateFilePath = "")
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

            if (string.IsNullOrEmpty(templateFilePath))
            {
                _sheet = string.IsNullOrEmpty(SheetName)
                    ? _workbook.CreateSheet()
                    : _workbook.CreateSheet(SheetName);
            }
            else
            {
                _workbook = WorkbookFactory.Create(templateFilePath);
                _sheet = _workbook.GetSheetAt(0);
            }


            var extractor = new ExileExcelExtractor<T>(_sheet);
            if (!string.IsNullOrEmpty(TitleText))
                extractor.DocumentMeta.TitleText = TitleText;

            extractor.DocumentMeta.IsUseTemplate = !string.IsNullOrEmpty(templateFilePath);

            if (dataList is IList)
            {

            }
            extractor.FillContent(dataList);

            _workbook.Write(stream);
        }

        private void WordWriteStream(IList<T> dataList, Stream stream, string templateFilePath = "")
        {
            using (var ms = new MemoryStream())
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
