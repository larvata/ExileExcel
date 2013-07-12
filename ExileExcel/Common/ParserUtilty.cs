namespace ExileExcel.Common
{
    using ExileExcel.Attribute;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    public static class ParserUtilty
    {
        /// <summary>
        /// Aquire HeaderText--Property mapping 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="headerList"></param>
        /// <returns></returns>
        internal static ExileMatchResult GetTypeMatched<T>(Dictionary<int, string> headerList)
        {
            var retVal = GetTypeMatched<T>();
            var currentKeyPair=new Dictionary<string, string>();

            var plist = currentKeyPair.Where(p => !headerList.ContainsValue(p.Value));
            // all matched
            if (!plist.Any())
            {
                return retVal;
            }

            return null;
        }

        /// <summary>
        /// Aquire HeaderText--Property mapping 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        internal static ExileMatchResult GetTypeMatched<T>()
        {
            //var retVal = new ExileMatchResult();
            //var currentKeyPair = new Dictionary<string, string>();

            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.GetCustomAttributes(typeof(ExilePropertyAttribute), true).Any());

            if (!properties.Any()) throw new Exception("MATCHING TYPE WITHOUT ATTRIBUTE ExilePropertyAttribute");

            // aquire class description
            var exiliableAttr =
                typeof(T).GetCustomAttributes(typeof(ExiliableAttribute), true).First() as ExiliableAttribute;

            var retVal = new ExileMatchResult
            {
                MatchedType = typeof (T),
                RowHeight = exiliableAttr.RowHeight,
                FontHeight = exiliableAttr.FontHeight,
                SheetName = exiliableAttr.SheetName,
                TitleText = exiliableAttr.TitleText,
                Visibility = exiliableAttr.Visibility
            };

            foreach (var property in properties)
            {
                var attr = property.GetCustomAttributes(typeof(ExilePropertyAttribute), true).First() as ExilePropertyAttribute;
                retVal.Headers.Add(new ExileHeader
                {
                    BuiltinFormat = attr.ColumnBulitinDataFormat,
                    CustomDataFormat = attr.ColumnCustomDataFormat,
                    DisplaySequence = attr.HeaderTextSequence,
                    PropertyDescription = attr.HeaderText,
                    PropertyName = property.Name, 
                    ColumnType = attr.ColumnType
                    
                });
            }


            return retVal;
        }


        //public static Dictionary<string, string> GetNameAttributePair(IEnumerable<Type> list)
        //{

        //    var retVal = new Dictionary<string, string>();

        //    foreach (var type in list)
        //    {
        //        var properties = type.GetProperties(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance)
        //            .Where(p=>p.GetCustomAttributes(typeof(ExilePropertyAttribute),true).Any());

        //        if (!properties.Any()) continue;

        //        foreach (var property in properties)
        //        {
        //            var attr = property.GetCustomAttributes(typeof (ExilePropertyAttribute), true).First() as ExilePropertyAttribute;
        //            retVal.Add(property.Name,attr.HeaderText);
        //        }
        //    }


        //    return retVal;
        //}

        public static IEnumerable<Type> GetTypesByAttribute(Type attributeType)
        {
            foreach (Type type in AppDomain.CurrentDomain.GetAssemblies().SelectMany(a => a.GetTypes()))
            {
                if (type.GetCustomAttributes(attributeType, true).Length > 0)
                {
                    yield return type;
                }
            }
        }

        /// <summary>
        /// get attribute value by attribute name
        /// </summary>
        /// <param name="src"></param>
        /// <param name="propName"></param>
        /// <returns></returns>
        public static ExilePropertyAttribute GetTypeAttribute(Type src, string propName)
        {
            return src.GetProperty(propName).GetCustomAttributes(typeof(ExilePropertyAttribute), true).First() as ExilePropertyAttribute;
        }

        public static object GetPropValue(object src, string propName)
        {
            return src.GetType().GetProperty(propName).GetValue(src, null);
        }

        public static Type GetPropType(object src, string propName)
        {
            return src.GetType().GetProperty(propName).PropertyType;
        }
    }
}
