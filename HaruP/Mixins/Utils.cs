using System;
using System.Reflection;
using System.Runtime.CompilerServices;

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
            catch
            {
            }

            return ret;
        }

        public static Type GetPropType(object src, string propName)
        {
            return src.GetType().GetProperty(propName).PropertyType;
        }
    }
}
