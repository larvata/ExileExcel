namespace ExileExcel.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    internal static class ReflectiveEnumerator
    {
        internal static IEnumerable<T> GetEnumerableOfType<T>(params object[] constructorArgs) where T : class
        {
            //var subclasses =
            //    AppDomain.CurrentDomain.GetAssemblies()
            //             .SelectMany(assembly => assembly.GetTypes(), (assembly, type) => new {assembly, type})
            //             .Where(@t => @t.type.IsSubclassOf(typeof (T)))
            //             .Select(@t => @t.type);

            //foreach (Type type in
            //    Assembly.GetAssembly(typeof (T)).GetTypes())
            //            //.Where(myType => myType.IsClass && !myType.IsAbstract && myType.IsSubclassOf(typeof (T))))
            //{
            //    objects.Add((T) Activator.CreateInstance(type, constructorArgs));
            //}
            //objects.Sort();


            // miao le ge mi de!
            return AppDomain.CurrentDomain.GetAssemblies().SelectMany(a => a.GetTypes(), (a, type) => new {a, type}).Where(t => t.type.IsSubclassOf(typeof (T))).Select(t => t.type).Select(type => (T) Activator.CreateInstance(type, constructorArgs)).ToList();
        }
    }
}