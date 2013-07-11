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

            if (retVal.HeaderKeyPair.Count == headerList.Count)
            {
                var plist = currentKeyPair.Where(p => !headerList.ContainsValue(p.Value));
                // all matched
                if (!plist.Any())
                {
                    //retVal.HeaderKeyPair = currentKeyPair;
                    return retVal;
                }
                
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
            var retVal = new ExileMatchResult();
            //var currentKeyPair = new Dictionary<string, string>();

            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.GetCustomAttributes(typeof(ExilePropertyAttribute), true).Any());

            if (!properties.Any()) throw new Exception("MATCHING TYPE WITH ATTRIBUTE ExilePropertyAttribute");

            foreach (var property in properties)
            {
                var attr = property.GetCustomAttributes(typeof(ExilePropertyAttribute), true).First() as ExilePropertyAttribute;
                retVal.HeaderKeyPair.Add(property.Name, attr.HeaderText);
            }

            // aquire class description
            retVal.TypeDescription = (typeof(T).GetCustomAttributes(typeof(ExiliableAttribute), true).First() as ExiliableAttribute).Description;

            // aquire type of class
            retVal.MatchedType = typeof (T);
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
