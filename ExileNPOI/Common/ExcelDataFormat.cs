using System;
using System.ComponentModel;
using NPOI.SS.UserModel;

namespace ExileNPOI.Common
{
    public enum ExcelDataFormatEnum
    {
        [Description("General")] General = 0, //0 General
        [Description("0")] NumberInteger, //1 0
        Number2DecimalPlace, //2 0.00
        NumberIntegerWithComma, //3 #,##0
        Number2DecimalPlaceWithComma, //4 #,##0.00
        CurrencyInteger, //5 "$"#,##0_);("$"#,##0)
        CurrencyIntegerWithRedAlert, //6 "$"#,##0_);[Red]("$"#,##0)
        Currency2DecimalPlace, //7 "$"#,##0.00_);("$"#,##0.00)
        Currency2DecimalPlaceWithRedAlert, //8 "$"#,##0.00_);[Red]("$"#,##0.00) 
        PercentInteger, //9 0%
        Percent2DecimalPlace, //10 0.00%
        Scientific2Digits, //11 0.00E+00
        Fraction, //12 # ?/? 
        Fraction2Digits, //13 # ??/?? 
        DateSlash, //14 M/D/YY
        DateDash, //15 D-MMM-YY
        DataDashNoYear, //16 D-MMM 
        DataDashNoDay, //17 MMM-YY
        Time12H, //18 Time h:mm AM/PM
        Time12HWithSeconds, //19 h:mm:ss AM/PM
        Time12HWithoutTag, //20 h:mm
        TimeWithSecond, //21 h:mm:ss
        DateTime, //22 M/D/YY h:mm
        //23~36 reserved
        Account = 37, //37 _(#,##0_);(#,##0)
        AccountWithRedAlert, //38 _(#,##0_);[Red](#,##0)
        Account2DecimalPlace, //39 _(#,##0.00_);(#,##0.00)
        Account2DecimalPlaceWithRedAlert, //40 _(#,##0.00_);[Red](#,##0.00)
        //41~44 Currency
        Time24H = 45, //45 mm:ss
        TimeCounts, //46 [h]:mm:ss
        TimeCountsWithDegree, //47 mm:ss.0
        Scientific, //48 ##0.0E+0
        Text //49 @
    }

    public static class ExcelDataFormat
    {
        /// <summary>
        /// Default value of DataTime format.Initialized as "YYYY/MM/DD HH:mm:ss"
        /// </summary>
        public static string DefaultDateTimeFormat { get; set; }

        /// <summary>
        /// Decimal Place count.Initialized as 2 
        /// </summary>
        public static int DefalultDecimalPlace { get; set; }


        static ExcelDataFormat()
        {
            DefaultDateTimeFormat = "YYYY/MM/DD HH:mm:ss";
            DefalultDecimalPlace = 2;
        }


        /// <summary>
        /// Convert .Net type to Excel data format string
        /// </summary>
        /// <param name="type">.Net type</param>
        /// <returns>Excel data format (eg. float returns "0.00" with default parameter)</returns>
        public static string Conver2ExcelDataFormat(Type type)
        {
            //IDataFormat retVal=workbook.CreateDataFormat();
            var formatString = string.Empty;

            if (type == typeof (int))
            {
                //retVal=(int)NPOIDataFormatEnum.NumberInteger;
                formatString = BuiltinFormats.GetBuiltinFormat((int) ExcelDataFormatEnum.NumberInteger);
            }
            else if (type == typeof (string))
            {
                //retVal=(int)NPOIDataFormatEnum.Text;
                formatString = BuiltinFormats.GetBuiltinFormat((int) ExcelDataFormatEnum.Text);
            }
            else if (type == typeof (Single) || type == typeof (float) || type == typeof (double))
            {
                switch (DefalultDecimalPlace)
                {
                    case 0:
                        //retVal=(int) NPOIDataFormatEnum.NumberInteger;
                        formatString = BuiltinFormats.GetBuiltinFormat((int) ExcelDataFormatEnum.NumberInteger);
                        break;
                    case 2:
                        //retVal=(int) NPOIDataFormatEnum.Number2DecimalPlace;
                        formatString = BuiltinFormats.GetBuiltinFormat((int) ExcelDataFormatEnum.Number2DecimalPlace);
                        break;
                    default:
                    {
                        const string formatTemplate = "0.";
                        formatString = formatTemplate.PadRight(DefalultDecimalPlace, '0');
                        //retVal=formatString;
                    }
                        break;
                }
            }
            else if (type == typeof (DateTime))
            {

                //retVal=dataFormat.GetFormat(DefaultDateTimeFormat);
                formatString = DefaultDateTimeFormat;
            }

            return formatString;

        }

    }

}
