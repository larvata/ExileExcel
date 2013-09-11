using System;
using System.ComponentModel;

namespace ExileNPOI.Attribute
{
    [Description("Header Cell Style Define")]
    [AttributeUsage(AttributeTargets.Property)]
    public class ExileHeaderGeneralAttribute:System.Attribute
    {

        /// <summary>
        /// Excel sheet header text
        /// </summary>
        public string HeaderText { get; set; }
        public int HeaderSequence { get; set; }
    }
}
