using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using HaruP.Parser;
using NPOI.SS.UserModel;

namespace HaruP.Common
{
    internal static class Utils
    {
        internal static bool CheckIfAnonymousType(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            // HACK: The only way to detect anonymous types right now.
            return System.Attribute.IsDefined(type, typeof(CompilerGeneratedAttribute), false)
                && type.IsGenericType && type.Name.Contains("AnonymousType")
                && (type.Name.StartsWith("<>") || type.Name.StartsWith("VB$"))
                && (type.Attributes & TypeAttributes.NotPublic) == TypeAttributes.NotPublic;
        }

        internal static object GetPropValue(object src, string propName)
        {
            object ret = null;
            try
            {
                ret =src.GetType().GetProperty(propName).GetValue(src, null);
            }
            catch(Exception ex)
            {
                throw ex;
            }

            return ret;
        }

        internal static IEnumerable<T> GetEnumerableOfType<T>(params object[] constructorArgs) where T : class
        {
            var ret = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(a => a.GetTypes(), (a, type) => new { a, type })
                .Where(t => t.type.GetInterface("IBaseSheetData") != null)
                .Select(t => t.type)
                .Select(type => (T)Activator.CreateInstance(type, constructorArgs))
                .ToList();
            return ret;
        }

        public static Dictionary<string, string> GetNameAttributePair(this IBaseSheetData baseExileData)
        {
            var retVal = new Dictionary<string, string>();

            var properties = baseExileData.GetType().GetProperties(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);

            foreach (var propertyInfo in properties)
            {
                var attrList = propertyInfo.GetCustomAttributes(typeof(SheetDataAttribute), true);

                if (attrList.Length == 0) continue;
                var attr = attrList[0] as SheetDataAttribute;
                if (attr != null) retVal.Add(propertyInfo.Name, attr.HeaderText);
            }

            return retVal;
        }

        public static string GetClassDescription(this IBaseSheetData baseExileData)
        {
            var attrs =
                System.Attribute.GetCustomAttributes(baseExileData.GetType()).OfType<DescriptionAttribute>().FirstOrDefault();
            return attrs != null ? attrs.Description : string.Empty;
        }

        public static bool IsKeyPairMatch(this IBaseSheetData baseSheetData, Dictionary<int, string> inputKeyPair)
        {
            var thisKeyPair = baseSheetData.GetNameAttributePair();
            if (inputKeyPair.Count != thisKeyPair.Count)
            {
                return false;
            }

            return thisKeyPair.All(
                inP => inputKeyPair.Any(p => p.Value.Equals(inP.Value)));
        }

        internal static Type GetPropType(object src, string propName)
        {
            return src.GetType().GetProperty(propName).PropertyType;
        }

        internal static IRow CellFormulaShift(this IRow row, int offset)
        {
            const string formulaExp = @"([A-Za-z]+)([\d]+)";
            foreach (ICell cell in row)
            {
                if (cell.CellType.Equals(CellType.Formula) && !string.IsNullOrEmpty(cell.CellFormula))
                {
                    var newFormulaString = Regex.Replace(cell.CellFormula, formulaExp,
                        m => m.Groups[1].Value + (Convert.ToInt16(m.Groups[2].Value) + offset).ToString());
                    cell.CellFormula = newFormulaString;
                }
            }
            return row;
        }

        internal static IRow CopyRowToAdvance(this IRow row, int rowNum, RowHeight rowHeight)
        {
            var newRow = row.CopyRowTo(rowNum);
            switch (rowHeight)
            {
                case RowHeight.Auto:
                    break;
                case RowHeight.Inherit:
                    newRow.Height = row.Height;
                    newRow.HeightInPoints = row.HeightInPoints;
                    break;
                default:
                    throw new ArgumentOutOfRangeException("rowHeight");
            }
            return row;
        }

    }
}
