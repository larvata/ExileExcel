using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using HaruP.Common;
using NPOI.SS.UserModel;

namespace HaruP.Mixins
{
    internal static class Utils
    {
        public static bool CheckIfAnonymousType(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            // HACK: The only way to detect anonymous types right now.
            return System.Attribute.IsDefined(type, typeof(CompilerGeneratedAttribute), false)
                && type.IsGenericType && type.Name.Contains("AnonymousType")
                && (type.Name.StartsWith("<>") || type.Name.StartsWith("VB$"))
                && (type.Attributes & TypeAttributes.NotPublic) == TypeAttributes.NotPublic;
        }

        public static object GetPropValue(object src, string propName)
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

        public static Type GetPropType(object src, string propName)
        {
            return src.GetType().GetProperty(propName).PropertyType;
        }

        public static IRow CellFormulaShift(this IRow row,int offset)
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

        public static IRow CopyRowToAdvance(this IRow row,int rowNum, RowHeight rowHeight)
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
