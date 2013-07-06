using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExileExcel.Attribute;

namespace ExileExcel.Common
{
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
            ExileMatchResult retVal = null;
            var currentKeyPair = new Dictionary<string, string>();

            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.GetCustomAttributes(typeof(ExilePropertyAttribute), true).Any());

            if (!properties.Any()) throw new Exception("MATCHING TYPE WITH ATTRIBUTE ExilePropertyAttribute");

            foreach (var property in properties)
            {
                var attr = property.GetCustomAttributes(typeof(ExilePropertyAttribute), true).First() as ExilePropertyAttribute;
                currentKeyPair.Add(property.Name, attr.HeaderText);
            }
            if (currentKeyPair.Count == headerList.Count)
            {
                var plist = currentKeyPair.Where(p => !headerList.ContainsValue(p.Value));
                // matched
                if (!plist.Any())
                {
                    retVal = (new ExileMatchResult
                        {
                            HeaderKeyPair = currentKeyPair,
                            MatchedType = typeof(T)
                        });
                }
            }

            // aquire class description
            retVal.TypeDescription = (typeof(T).GetCustomAttributes(typeof(ExiliableAttribute), true).First() as ExiliableAttribute).Description;

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


    }
}
