using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace ExileExcel
{
    public class BaseExileData 
    {
        public int ForeignHashCode { get; set; }

        public string Description
        {
            get
            {
                return GetClassDescription();
            }
        }


        public Dictionary<string, string> GetNameAttributePair()
        {
            var retVal = new Dictionary<string, string>();

            var properties = this.GetType().GetProperties(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);

            foreach (var propertyInfo in properties)
            {
                var attrList = propertyInfo.GetCustomAttributes(typeof(ExileDataAttribute), true);

                if (attrList.Length == 0) continue;
                var attr = attrList[0] as ExileDataAttribute;
                if (attr != null) retVal.Add(propertyInfo.Name, attr.HeaderText);
            }

            return retVal;
        }

        public string GetClassDescription()
        {

            var attrs =
                System.Attribute.GetCustomAttributes(this.GetType()).OfType<DescriptionAttribute>().FirstOrDefault();
            return attrs != null ? attrs.Description : string.Empty;
        }

        public bool IsKeyPairMatch(Dictionary<int, string> inputKeyPair)
        {
            var thisKeyPair = GetNameAttributePair();
            if (inputKeyPair.Count != thisKeyPair.Count)
            {
                return false;
            }

            return thisKeyPair.All(
                inP => inputKeyPair.Any(p =>
                    p.Value.Equals(inP.Value)));
        }
    }
}
